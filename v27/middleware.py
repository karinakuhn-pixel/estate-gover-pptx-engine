"""Operation correlation for every request reaching the isolated V2.7 path."""

from uuid import uuid4

from starlette.middleware.base import BaseHTTPMiddleware

from .audit import AuditLogger


class V27OperationIdMiddleware(BaseHTTPMiddleware):
    def __init__(self, app, audit_logger: AuditLogger):
        super().__init__(app)
        self.audit_logger = audit_logger

    async def dispatch(self, request, call_next):
        if not request.url.path.startswith("/v2.7"):
            return await call_next(request)

        operation_id = str(uuid4())
        request.state.operation_id = operation_id
        self.audit_logger.record({
            "event": "request_started",
            "operation_id": operation_id,
            "path": request.url.path,
            "method": request.method,
            "publicacao_automatica": False,
            "status_arquivo": "RASCUNHO",
        })

        try:
            response = await call_next(request)
        except Exception as exc:
            self.audit_logger.record({
                "event": "request_failed",
                "operation_id": operation_id,
                "result": "erro",
                "error_type": type(exc).__name__,
                "publicacao_automatica": False,
                "status_arquivo": "RASCUNHO",
            })
            raise

        response.headers["X-Operation-ID"] = operation_id
        self.audit_logger.record({
            "event": "request_completed",
            "operation_id": operation_id,
            "result": "sucesso" if response.status_code < 400 else "erro",
            "http_status": response.status_code,
            "publicacao_automatica": False,
            "status_arquivo": "RASCUNHO",
        })
        return response

