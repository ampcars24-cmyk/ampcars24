# Google Drive export relay

This relay lets the authenticated portal server load the latest Excel exports directly from Drive. The browser continues to call the portal's `/api/data` endpoint. The server validates OIDC, calls this relay with a server-only secret, and filters partner rows before returning data to the browser.

## Configure Apps Script

1. Create or open an Apps Script project and replace its code with `Code.gs`.
2. In **Project Settings → Script properties**, add:
   - `RELAY_SHARED_SECRET`: a long random secret. Keep it private.
   - `EXPORT_FOLDER_ID`: the Drive folder ID containing the exports. The folder previously shared for this portal was `1YEu-gaQICC7Fp2Hc5dEXGRDP972x8jVH`; confirm that it contains both export files before using it.
   - `ADMIN_EXPORT_FILENAME`: `Sampled X Orignal.xlsx` (optional; this is the default).
   - `PARTNER_EXPORT_FILENAME`: `Report Level.xlsx` (optional; this is the default).
3. Add the **Drive API** advanced service (v3). If the Apps Script project uses a standard Google Cloud project, enable the Drive API there too.
4. Deploy as a Web App that executes as the file owner. The endpoint can be reachable publicly because data actions require the server-only secret; do not put that secret in the browser or in `index.html`.
5. Copy the `/exec` URL and set `DRIVE_RELAY_URL` and the same `DRIVE_RELAY_SECRET` in the portal server's environment. Leave both unset to keep the existing local export fallback.

Keep the current daily files under the configured names (`Sampled X Orignal.xlsx` and `Report Level.xlsx`). The relay prefers those exact files for the normal refresh. To retain monthly dumps, save copies using a month/year suffix, for example `Sampled X Orignal - Nov 2026.xlsx` and `Report Level - Nov 2026.xlsx`. The portal’s month selector lists available archive periods; if the unqualified daily file is absent, the relay chooses the newest month/year file. To request a specific archived period through the authenticated server API, use `/api/data?period=Nov-2026`.

The relay converts each selected XLSX to a temporary Google Sheet for reading and deletes that temporary conversion after the request.

## Request shape

The portal server sends a POST body like:

```json
{"action":"getData","dataset":"partner","token":"server-only-secret"}
```

`dataset` is `admin` or `partner`. The relay returns the corresponding workbook rows; the portal server performs the OIDC-based partner filter. The Config-sheet read/write actions and file upload action also require the same secret.

