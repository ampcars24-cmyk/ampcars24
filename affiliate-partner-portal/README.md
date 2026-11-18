# CARS24 Affiliate Partner Portal

This folder restores the existing AFF Partner Portal page and replaces its Google Sheet refresh with the two scheduled Apex Excel exports. The Sell Reviews app is separate and is not included in this portal folder.

## Data refresh

The Node server reads the latest files on every authenticated `/api/data` request:

- Admins: `Sampled X Orignal.xlsx`
- Partners: `Report Level.xlsx`, filtered on the server to the partner linked to the verified OIDC subject

On the current Windows setup the Google Drive for desktop folder is `C:\Users\40980\Downloads\AFF Partner Portal`. Replace both workbooks in that synced folder after the daily export. No workbook data is stored in this repository.

## Deployment requirements

GitHub Pages cannot read a private Drive for desktop folder or verify OIDC access tokens. Run this Node app on a host that can read the synchronized export directory, or configure another private file mount for the same two filenames.

1. Install Node.js 20 or newer and run `npm install` in this folder.
2. Copy `.env.example` to `.env` and set the actual OIDC issuer, audience, JWKS URL, authority, client ID, API scope, verified admin role claim, and subject-to-partner map.
3. In the identity provider, register redirect URI `https://<portal-host>/auth/callback` and allow the portal origin.
4. Set `EXPORT_DATA_DIR` to the folder containing the two current workbooks.
5. Run `npm start` behind HTTPS.

Until real OIDC values and subject mappings are configured, sign-in stays disabled. Do not put passwords, OIDC client secrets, or exported workbooks in GitHub.

