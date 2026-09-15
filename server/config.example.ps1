# Copy to config.local.ps1 and fill in your own values.
# Load in PowerShell with: . ./config.local.ps1
$env:ADMIN_EMAIL = "admin@example.com"
$env:SMTP_HOST = "smtp.example.com"
$env:SMTP_PORT = "465"
$env:SMTP_USER = "sender@example.com"
$env:SMTP_PASS = "REPLACE_WITH_SMTP_APP_PASSWORD"
$env:DRM_SHARED_SECRET = "REPLACE_WITH_SHARED_SECRET"
