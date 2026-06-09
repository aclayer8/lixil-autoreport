import os

APP_HOST = os.getenv("LIXIL_APP_HOST", "127.0.0.1")
APP_PORT = int(os.getenv("LIXIL_APP_PORT", "5000"))
AUTO_OPEN_BROWSER = os.getenv("LIXIL_AUTO_OPEN_BROWSER", "1").lower() in {"1", "true", "yes", "on"}
SHOW_ERROR_TRACE = os.getenv("LIXIL_SHOW_ERROR_TRACE", "0").lower() in {"1", "true", "yes", "on"}
MAX_UPLOAD_MB = int(os.getenv("LIXIL_MAX_UPLOAD_MB", "50"))
