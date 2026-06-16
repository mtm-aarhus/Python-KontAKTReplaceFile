# Python-KontAKTReplaceFile

Replaces a document's file in SharePoint with a manually-redacted version uploaded by the caseworker, for the **KontAKT** aktindsigt (FOI request) system.

Some documents can't be converted to PDF (e.g. spreadsheets, e-mails, unusual formats), so they can't be redacted in KontAKT's in-browser editor. Instead the caseworker downloads the original, redacts it in a suitable program, and uploads the redacted file back to KontAKT (**"Redigér manuelt"**). This robot pushes that upload to SharePoint, since only it can write there.

## What it does

For one document:

1. Fetches the current SharePoint URL and the target filename from KontAKT (the target name is the existing one, but with the uploaded file's extension).
2. Downloads the staged (browser-uploaded) redacted file from KontAKT.
3. Uploads it to the same SharePoint folder, **replacing the original file**.
4. If the redacted upload has a different name/extension (e.g. a redacted `.xlsx` exported as `.pdf`), deletes the old file so the folder isn't left with both.
5. Reports back to KontAKT with the new URL, filename, hash and size; the document is marked `redacted`.

The un-redacted source still lives in GO/Nova — only the SharePoint copy is replaced.

## Input (one document)

| Field | Meaning |
|-------|---------|
| `kontakt_case_id` | KontAKT case id |
| `doc_id` | KontAKT document id |

The uploaded file itself and the target filename are fetched from KontAKT (files don't fit in the queue payload).

## Output

The uploaded file replaces the original in SharePoint, plus a callback to KontAKT:

```json
{"ok": true, "sharepoint_url": "https://…", "filename": "…", "sha256": "…", "file_size_bytes": 12345}
```

The new `sha256` lets KontAKT bust its cache so "Åbn" shows the redacted file. On failure it reports `{"ok": false, "note": "…"}` and the document is marked `error` so the caseworker can re-upload.

## Required configuration

- Constant `KontAKTSharePoint` — SharePoint site URL
- Credential `SharePointCert` — username = certificate thumbprint, password = certificate path
- Credential `SharePointAPI` — username = tenant, password = client id
- Credential `KontAKTAPI` — username = base URL, password = API key

## Dependencies

The shared [`oomtm`](https://github.com/mtm-aarhus/oomtm) library (`sharepoint`) for the upload/delete. No PDF or OCR work — the file is already redacted by a human.
