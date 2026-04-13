import os
import sys
from .tools import mcp
from .auth import set_external_bearer


def main() -> None:
    if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
        print(
            "Error: MICROSOFT_MCP_CLIENT_ID environment variable is required",
            file=sys.stderr,
        )
        sys.exit(1)

    external_bearer_mode = os.getenv("EXTERNAL_BEARER_MODE", "").lower() == "true"
    transport = os.getenv("MCP_TRANSPORT", "stdio")
    host = os.getenv("MCP_HOST", "0.0.0.0")
    port = int(os.getenv("PORT", os.getenv("MCP_PORT", "8000")))

    if external_bearer_mode:
        print("[INFO] External bearer mode enabled")
        print("[INFO] Accepting Microsoft access tokens from Authorization header")

        from starlette.middleware.base import BaseHTTPMiddleware
        from starlette.requests import Request

        class BearerMiddleware(BaseHTTPMiddleware):
            async def dispatch(self, request: Request, call_next):
                auth_header = request.headers.get("authorization", "")
                if auth_header.lower().startswith("bearer "):
                    set_external_bearer(auth_header[7:])
                return await call_next(request)

        mcp.settings.host = host
        mcp.settings.port = port

        app = mcp.get_app(transport="streamable-http")
        app.add_middleware(BearerMiddleware)

        import uvicorn
        print(f"[INFO] Starting streamable-http on {host}:{port}/mcp")
        uvicorn.run(app, host=host, port=port)
    elif transport == "streamable-http":
        mcp.run(transport="streamable-http", host=host, port=port)
    else:
        mcp.run()


if __name__ == "__main__":
    main()
