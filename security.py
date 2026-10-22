import os
from fastapi import Depends, Header, HTTPException, status

API_TOKEN = os.environ.get("DOC_PDF_API_TOKEN")
if API_TOKEN is None:
    raise RuntimeError("Env-Var DOC_PDF_API_TOKEN fehlt!")

async def api_key_auth(x_api_key: str | None = Header(default=None)):
    if x_api_key != API_TOKEN:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Ungültiger oder fehlender API-Token",
            headers={"WWW-Authenticate": "X-API-Key"},
        )
