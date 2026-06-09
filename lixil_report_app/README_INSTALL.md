# LIXIL Report App

Internal web app for importing Cisco call manager/CUIC/CAR raw reports, summarizing the data, showing charts/tables, and exporting Excel summaries.

## Install

1. Install Python 3.10 or newer.
2. Run `setup.bat` once. This creates `.venv` and installs dependencies from `requirements.txt`.

## Run on a Local Desktop

Run `run_local.bat`.

Default URL: `http://127.0.0.1:5000`

## Run on an Internal Customer Server

Run `run_server.bat`.

Default bind: `0.0.0.0:5000`

Users on the same network can open:

`http://SERVER_NAME:5000`

If Windows Firewall blocks access, allow inbound TCP port `5000` for the server.

## Environment Variables

- `LIXIL_APP_HOST`: host to bind when running `app.py`; default `127.0.0.1`
- `LIXIL_APP_PORT`: port; default `5000`
- `LIXIL_AUTO_OPEN_BROWSER`: `1` to auto-open browser for desktop mode, `0` for server mode
- `LIXIL_SHOW_ERROR_TRACE`: `1` to return debug trace in API responses, `0` for customer/server mode
- `LIXIL_MAX_UPLOAD_MB`: upload limit in MB; default `50`

## Current Report Status

- 1.1 to 1.7: implemented and smoke-tested with current samples
- 1.8 to 1.10: implementation exists but should be validated with customer samples/spec
- 1.11: implemented and smoke-tested with current sample

## Notes Before Delivery

- Keep raw/customer sample files outside the app folder when packaging for production.
- The current UI still references Bootstrap and Chart.js from CDN. For fully offline customer environments, vendor those files into `static/vendor` and update `templates/index.html`.
- Do not enable `LIXIL_SHOW_ERROR_TRACE=1` on the customer server.
