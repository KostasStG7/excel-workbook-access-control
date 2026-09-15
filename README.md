# Excel Workbook Access Control

A personal Excel VBA project for workbook access control and administrator-approved enrollment, supported by a Python Flask notification server.

## Overview

The main application runs inside Excel. VBA handles the enrollment form, local access checks, worksheet visibility, workbook protection and enrollment logs. The supporting Flask server receives challenge notifications from the workbook over HTTP/JSON and forwards them to an administrator by email.

## Features

- Challenge-response enrollment with a time limit on VBA response validation.
- Checks against configured Windows usernames and workbook folder paths.
- Administrator tools for processing enrollment requests and viewing logs.
- Worksheet visibility and protection changes during open, save and close events.
- JSON webhook notifications and SMTP email delivery.
- Optional Cloudflare Tunnel connection to the local notification server.

The function named `DeviceName` currently uses the Windows username; this is not a hardware device identity.

## Repository contents

| Path | Purpose |
| --- | --- |
| `vba/DRM_Core.bas` | Main access control and enrollment logic |
| `vba/EnrollForm.frm` and `.frx` | Enrollment form code and its required binary form resources |
| `vba/Module1.bas` | Initial workbook structure setup |
| `vba/Module2.bas` | Workbook sealing procedure |
| `vba/ThisWorkbook.cls` | Workbook event handler source |
| `vba/OwnerToolsWorksheet.cls` | Administrator worksheet event handler source |
| `server/server.py` | Supporting Flask notification server |
| `server/config.example.ps1` | Example server environment configuration |

## Technology

Excel VBA, Python, Flask, HTTP/JSON, SMTP and optional Cloudflare Tunnel.

## Configuration and local evaluation

This repository contains source exports, not a ready-to-run workbook. The original `.xlsm`, its private configuration, existing enrollment records and deployment credentials are excluded.

### Excel VBA

Use desktop Excel for Windows and a disposable macro-enabled workbook for evaluation. Keep a backup of any workbook you modify.

1. Import the `.bas` modules and `EnrollForm.frm` in the VBA editor. Keep `EnrollForm.frx` beside the `.frm` when importing.
2. Set the configuration placeholders before running the code. Use one consistent sheet protection password wherever `REPLACE_WITH_SHEET_PWD` appears. Set an administrator password and a webhook shared secret. These example placeholders are not real credentials.
3. Run `InitDRMStructure` to initialize the supporting sheets.
4. Copy the procedure code from `ThisWorkbook.cls` into the workbook's existing `ThisWorkbook` code module, and from `OwnerToolsWorksheet.cls` into the `OWNER_TOOLS` worksheet module. Do not paste the export headers (`VERSION`, `BEGIN`, `END`, or `Attribute` lines), and do not import these as ordinary class modules.
5. Configure authorized usernames, folder paths and the workbook secret in `DRM_CFG`. Administrator and distributed workbook copies must use a matching workbook secret for the challenge-response workflow. The workbook secret is separate from the webhook shared secret.
6. Review `EXPIRY_DATE` and `WEBHOOK_URL` in `DRM_Core.bas` for your evaluation environment. The publication copy points to `http://127.0.0.1:5000/v1/challenge`; a remote endpoint requires your own configuration.
7. Compile the VBA project and test opening, enrollment, saving and closing in the disposable workbook before using it elsewhere.

The VBA exports retain their original Windows Greek encoding for import compatibility. Some original comments contain encoding artifacts.

### Supporting server

Use Python 3.10 or newer. In PowerShell, from `server/`:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item config.example.ps1 config.local.ps1
# Edit config.local.ps1 with your own values before continuing.
. ./config.local.ps1
.venv\Scripts\python.exe server.py
```

Set `DRM_SHARED_SECRET` to the same value as the VBA `SHARED_SECRET`. Use your provider's SMTP settings. Local configuration files and working Excel files must remain outside commits.

This repository includes the simpler notification server from the original project. Its `/v1/challenge` route checks the webhook key and sends email; enrollment response validation takes place in VBA. Unused verification helpers in the server are not evidence of endpoint validation. The separate experimental rate-limit server variant is not included.

## Scope and limitations

This is a personal prototype, not a production DRM or data encryption product. Excel protection and hidden sheets can be bypassed. The custom XOR/FNV-based routines are not modern cryptographic protection. Debug output and local enrollment logs can contain challenge details. The Flask development server is intended for local evaluation; deployment requires additional configuration and review.

## Publication preparation

Embedded credentials and deployment-specific addresses were replaced with placeholders. The Flask configuration reads credentials from environment variables. The original application logic has otherwise been retained. Python syntax and the publication file selection were checked; Excel/VBA execution, SMTP delivery and the complete enrollment flow have not been tested in this preparation environment.
