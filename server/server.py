import os, ssl, smtplib, hmac, hashlib, calendar, time
from email.mime.text import MIMEText
from flask import Flask, request, jsonify

app = Flask(__name__)

# ===== SMTP / EMAIL CONFIG =====
ADMIN_EMAIL = os.environ["ADMIN_EMAIL"]
SMTP_HOST   = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT   = int(os.environ.get("SMTP_PORT", "465"))       # SSL port
SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASS = os.environ["SMTP_PASS"]
DRM_SHARED_SECRET = os.environ["DRM_SHARED_SECRET"]
CH_TTL_SECONDS = int(os.environ.get("CH_TTL_SECONDS", "600"))  # 10'
RESP_LEN = int(os.environ.get("RESP_LEN", "8"))   # default 8 hex chars (όπως VBA RESP_LEN)



# ---- FNV1a 32bit implementation (to match VBA FNV1a32Hex) ----
def fnv1a32_hex(s: str) -> str:
    """
    Compute FNV-1a 32-bit and return uppercase 8-char hex string (like VBA FNV1a32Hex).
    """
    # FNV-1a parameters 32-bit
    FNV_OFFSET = 0x811c9dc5
    FNV_PRIME = 16777619
    h = FNV_OFFSET
    # Ensure we iterate raw bytes as VBA uses Asc(mid$(s,i,1)) & 0xFF
    b = s.encode('latin-1', errors='replace')  # latin-1 keeps bytes 0-255 stable
    for byte in b:
        h = h ^ byte
        # multiply 32-bit with overflow
        h = (h * FNV_PRIME) & 0xFFFFFFFF
    # convert to signed-like behavior compatible with VBA hex formatting
    # but we only need 8 hex digits uppercase
    return format(h & 0xFFFFFFFF, '08X')

def verify_response(ch_id: str, response: str, payload: str = None) -> bool:
    """
    Compute expected = Left$(FNV1a32Hex(DRM_SHARED_SECRET & "|" & payload), RESP_LEN)
    and compare to provided response (case-insensitive).
    Requires 'payload' provided (we use the one that client sent).
    """
    try:
        if payload is None:
            return False
        base = DRM_SHARED_SECRET + "|" + payload
        full_hex = fnv1a32_hex(base)     # e.g. "A1B2C3D4"
        expected = full_hex[:RESP_LEN].upper()
        # compare case-insensitive (VBA uses vbTextCompare in places)
        return hmac.compare_digest(expected.upper(), (str(response or "")).upper())
    except Exception:
        return False



def verify_hmac(secret, token, payload, ts, sig_hex):
    base = f"{token}|{payload}|{ts}".encode("utf-8")
    calc = hmac.new(secret.encode("utf-8"), base, hashlib.sha256).hexdigest()
    return hmac.compare_digest(calc, (sig_hex or "").lower())

def send_admin_mail(subject: str, body: str):
    msg = MIMEText(body, _charset="utf-8")
    msg["From"] = SMTP_USER
    msg["To"] = ADMIN_EMAIL
    msg["Subject"] = subject

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx, timeout=20) as s:
        s.set_debuglevel(0)                 # προαιρετικό: δείχνει SMTP διάλογο στο console
        s.login(SMTP_USER, SMTP_PASS)       # app password της Google
        s.sendmail(SMTP_USER, [ADMIN_EMAIL], msg.as_string())

# ===== ENDPOINT =====
@app.post("/v1/challenge")
def receive_challenge():
# quick header check
    key = request.headers.get("X-Webhook-Key", "")
    if not key or not hmac.compare_digest(key, DRM_SHARED_SECRET):
        return jsonify({"ok": False, "err": "bad-key"}), 403
    data = request.get_json(force=True, silent=True) or {}
    token   = str(data.get("token", ""))
    payload = str(data.get("payload", ""))
    ts_utc  = str(data.get("ts_utc", ""))
    wb_id   = str(data.get("workbook_id", ""))

    # φτιάξε το email σώμα
    subject = "Excel DRM Challenge"
    body = (
        f"Workbook: {wb_id}\n"
        f"Token:    {token}\n"
        f"Payload:  {payload}\n"
        f"UTC:      {ts_utc}\n"
        f"IP:       {request.remote_addr}\n"
    )

    try:
        send_admin_mail(subject, body)
    except Exception as e:
        print("Email send failed:", e)

    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
