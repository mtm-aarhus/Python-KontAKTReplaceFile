"""This module handles resetting the state of the computer so the robot can work with a clean slate.

For this robot the "state" is the SharePoint connection and the cached OO
credentials. ``open_all`` opens them and returns a :class:`Client`; ``reset``
re-opens them, so the queue framework can reconnect on a retry instead of doing
a fresh cert-auth handshake for every document. (The replacement files are
fetched from KontAKT, so no GO/Nova connection is needed.)
"""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection

from oomtm import sharepoint as sp


class Client:
    """Live SharePoint connection + cached KontAKT credentials, opened by
    ``open_all`` and reused across every queue element (a run can replace many
    documents' files and shares one cert-auth handshake)."""

    def __init__(self, orchestrator_connection: OrchestratorConnection):
        self.sp_ctx, self.sp_site_url = _build_sp_context(orchestrator_connection)
        kontakt = orchestrator_connection.get_credential("KontAKTAPI")
        self.kontakt_base = kontakt.username
        self.kontakt_key = kontakt.password


def reset(orchestrator_connection: OrchestratorConnection) -> Client:
    """Clean up, close/kill all programs, then (re)open the connections.

    Returns the freshly-opened :class:`Client` so the queue framework can reuse
    it across queue elements (and reconnect by calling ``reset`` again)."""
    orchestrator_connection.log_trace("Resetting.")
    clean_up(orchestrator_connection)
    close_all(orchestrator_connection)
    kill_all(orchestrator_connection)
    return open_all(orchestrator_connection)


def clean_up(orchestrator_connection: OrchestratorConnection) -> None:
    """Do any cleanup needed to leave a blank slate."""
    orchestrator_connection.log_trace("Doing cleanup.")


def close_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Gracefully close all applications used by the robot."""
    orchestrator_connection.log_trace("Closing all applications.")


def kill_all(orchestrator_connection: OrchestratorConnection) -> None:
    """Forcefully close all applications used by the robot."""
    orchestrator_connection.log_trace("Killing all applications.")


def open_all(orchestrator_connection: OrchestratorConnection) -> Client:
    """Open all connections used by the robot and return them as a :class:`Client`."""
    orchestrator_connection.log_trace("Opening SharePoint connection.")
    return Client(orchestrator_connection)


# ----- SharePoint context ----------------------------------------------------


def _build_sp_context(orchestrator_connection):
    cert = orchestrator_connection.get_credential("SharePointCert")  # user=thumbprint, pwd=cert_path
    api = orchestrator_connection.get_credential("SharePointAPI")     # user=tenant,    pwd=client_id
    raw = (orchestrator_connection.get_constant("KontAKTSharePoint").value or "").strip().rstrip("/")
    for suffix in ("/Delte dokumenter", "/Delte%20dokumenter"):
        if raw.lower().endswith(suffix.lower()):
            raw = raw[: -len(suffix)]
    site_url = raw.rstrip("/")
    ctx = sp.connect(
        site_url=site_url, tenant=api.username, client_id=api.password,
        thumbprint=cert.username, cert_path=cert.password,
    )
    return ctx, site_url
