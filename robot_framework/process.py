"""KontAKT file-replacement robot — swaps in a manually-redacted file.

Some documents can't be converted to PDF (e.g. .xlsx, .msg, odd formats), so
they can't be redacted in KontAKT's in-browser editor. Instead the caseworker
downloads the original, redacts it in a suitable program and uploads the
redacted version to KontAKT. This robot pushes that upload to SharePoint — the
web tier is SharePoint read-only, so the write has to happen here.

Queue-driven, one queue element per document. For a single document it:

  1. fetches the target filename + the current SharePoint URL from KontAKT,
  2. downloads the staged (browser-uploaded) file from KontAKT,
  3. uploads it to the same SharePoint folder, replacing the original,
  4. deletes the old file if the redacted upload has a different name/extension
     (e.g. the caseworker exported a redacted .xlsx as .pdf),
  5. reports back to KontAKT (new URL + filename + hash + size; the document is
     marked 'redacted' and KontAKT's cache is busted).

Queue payload (set by KontAKT's "Redigér manuelt" upload):
    {"kontakt_case_id": 11, "doc_id": 42}

OO config:
    Constant   KontAKTSharePoint      — SharePoint site URL
    Credential SharePointCert         — username = thumbprint, password = cert path
    Credential SharePointAPI          — username = tenant,     password = client id
    Credential KontAKTAPI             — username = base URL,    password = X-API-Key
"""
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
import hashlib
import json
import posixpath
import tempfile
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import requests

from robot_framework import reset
from oomtm import sharepoint as sp


def process(
    orchestrator_connection: OrchestratorConnection,
    queue_element: QueueElement | None = None,
    client: "reset.Client | None" = None,
) -> None:
    orchestrator_connection.log_trace("Running process.")
    if queue_element is None:
        raise RuntimeError("KontAKTReplaceFile is queue-driven; no queue_element given.")
    if client is None:  # e.g. a manual run outside the queue framework
        client = reset.open_all(orchestrator_connection)

    payload = json.loads(queue_element.data or "{}")
    case_id = int(payload["kontakt_case_id"])
    doc_id = int(payload["doc_id"])
    orchestrator_connection.log_info(f"Replace file case={case_id} doc={doc_id}")

    try:
        result = _replace(orchestrator_connection, client, case_id, doc_id)
    except Exception as exc:
        orchestrator_connection.log_info(f"Replace failed: {exc!r}")
        _callback(orchestrator_connection, client, case_id, doc_id, {"ok": False, "note": str(exc)[:500]})
        raise

    _callback(orchestrator_connection, client, case_id, doc_id, result)
    orchestrator_connection.log_info(f"Replace done doc={doc_id}: ok={result.get('ok')}")


def _replace(orchestrator_connection, client, case_id, doc_id):
    info = _fetch_info(client, case_id, doc_id)
    sharepoint_url = (info.get("sharepoint_url") or "").strip()
    target_filename = (info.get("target_filename") or "").strip()
    if not sharepoint_url:
        return {"ok": False, "note": "Dokumentet har ingen SharePoint-fil at erstatte."}
    if not target_filename:
        return {"ok": False, "note": "Mangler filnavn til den nye fil."}

    server_relative = unquote(urlparse(sharepoint_url).path)
    folder_path = posixpath.dirname(server_relative)
    old_filename = posixpath.basename(server_relative)

    with tempfile.TemporaryDirectory() as tmpdir:
        # Upload under the target name (= original name, possibly a new extension).
        dst = Path(tmpdir) / target_filename
        _download_staged(client, case_id, doc_id, dst)
        sha = _sha256_hex(dst)
        size = dst.stat().st_size
        sp.upload_file(client.sp_ctx, folder_path=folder_path, local_file=str(dst), overwrite=True)

    # Redacted upload renamed / changed format → remove the old file so the
    # folder doesn't keep both the original and the redacted version.
    if target_filename != old_filename:
        old_rel = posixpath.join(folder_path, old_filename)
        try:
            sp.delete_file(client.sp_ctx, old_rel)
        except Exception as exc:  # pylint: disable=broad-except
            orchestrator_connection.log_info(f"Could not delete old file {old_rel}: {exc!r}")

    new_url = sharepoint_url.rsplit("/", 1)[0] + "/" + quote(target_filename)
    return {"ok": True, "sharepoint_url": new_url, "filename": target_filename,
            "sha256": sha, "file_size_bytes": size}


def _sha256_hex(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# ----- KontAKT API -----------------------------------------------------------


def _fetch_info(client, case_id, doc_id):
    """GET the current SharePoint URL + the target filename for the new file."""
    r = requests.get(
        f"{client.kontakt_base}/api/v1/cases/{case_id}/documents/{doc_id}/replacement-info",
        headers={"X-API-Key": client.kontakt_key},
        timeout=30,
    )
    r.raise_for_status()
    return r.json()


def _download_staged(client, case_id, doc_id, dst: Path) -> None:
    """Stream the browser-uploaded redacted file from KontAKT to ``dst``."""
    r = requests.get(
        f"{client.kontakt_base}/api/v1/cases/{case_id}/documents/{doc_id}/replacement-file",
        headers={"X-API-Key": client.kontakt_key},
        timeout=120, stream=True,
    )
    r.raise_for_status()
    with open(dst, "wb") as fh:
        for chunk in r.iter_content(1 << 20):
            if chunk:
                fh.write(chunk)


def _callback(orchestrator_connection, client, case_id: int, doc_id: int, body: dict) -> None:
    try:
        requests.post(
            f"{client.kontakt_base}/api/v1/cases/{case_id}/documents/{doc_id}/file-replaced",
            headers={"X-API-Key": client.kontakt_key, "Content-Type": "application/json"},
            json=body, timeout=30,
        )
    except Exception as exc:  # pylint: disable=broad-except
        orchestrator_connection.log_info(f"Callback to KontAKT failed: {exc!r}")
