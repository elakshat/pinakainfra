# Backend Deployment

This portal is not a pure static HTML deployment. Host it on a Node-capable service such as Render, Railway, Fly.io, a VPS, or Firebase Cloud Run/Functions with Hosting rewrites.

The browser should call only same-origin routes like:

- `/api/portal/config`
- `/api/portal/unlock`
- `/api/portal/logout`
- `/api/master`
- `/api/fuel`
- `/api/weighbridge`
- `/api/user-charge`
- `/api/mrf`
- `/api/joining`
- `/api/maintenance`
- `/api/complaints`

The real Google Apps Script URLs must be configured as environment variables on the hosting server. They are intentionally not embedded in the browser files.

PIN validation happens on the backend. The browser receives portal app settings, but it does not receive the Master Sheet user/PIN list.

## Production Environment Variables

Set these on the hosting server:

- `MASTER_SCRIPT_URL`
- `FUEL_SCRIPT_URL`
- `WEIGHBRIDGE_SCRIPT_URL`
- `USER_CHARGE_SCRIPT_URL`
- `MRF_SCRIPT_URL`
- `JOINING_SCRIPT_URL`
- `MAINTENANCE_SCRIPT_URL`
- `COMPLAINTS_SCRIPT_URL`

Do not deploy this as static-only hosting, because `/api/...` routes require the Node backend. Static-only hosting will load the shell but app data and protected actions will fail.

Protected app API routes require the backend session cookie created by `/api/portal/unlock`.

## Complaints Google Sheets Backend

The complaints app includes a new Apps Script backend at:

`apps/scripts/complaints-script.gs`

Create or open the Google Sheet that should store complaints, then:

1. Open `Extensions > Apps Script`.
2. Paste the full contents of `apps/scripts/complaints-script.gs`.
3. Run `setupComplaintsSheet_` once from Apps Script to create the `Complaints` and `Settings` tabs.
4. Deploy as a Web App:
   - Execute as: `Me`
   - Who has access: `Anyone`
5. Copy the `/exec` URL into `COMPLAINTS_SCRIPT_URL` in `.env`.

Until `COMPLAINTS_SCRIPT_URL` is filled, the complaints app falls back to local demo storage.

## Start Command

Use:

```bash
npm start
```

The server reads `PORT` from the host automatically.

## Protected Files

The Node server blocks browser access to:

- `server.mjs`
- markdown/docs files
- import files
- `apps/scripts/`
