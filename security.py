import hmac
import os
from fastapi import Header, HTTPException, Request, status

API_TOKEN = os.environ.get("DOC_PDF_API_TOKEN", "")
if not API_TOKEN:
    raise RuntimeError("Env-Var DOC_PDF_API_TOKEN fehlt oder ist leer!")

# Routes reachable without an API token (everything else is protected)
PUBLIC_PATHS = {"/healthz"}

async def api_key_auth(request: Request, x_api_key: str | None = Header(default=None)):
    route = request.scope.get("route")
    if getattr(route, "path", None) in PUBLIC_PATHS:
        return
    # compare_digest runs in constant time, so response timing doesn't leak the token
    if x_api_key is None or not hmac.compare_digest(x_api_key.encode(), API_TOKEN.encode()):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Ungültiger oder fehlender API-Token",
            headers={"WWW-Authenticate": "X-API-Key"},
        )
