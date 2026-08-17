"""Governed V2.7 route with Phase-1 safe fallback and Phase-2 CAPA homologation."""

from __future__ import annotations

import hashlib
import json
import tempfile
from pathlib import Path

from fastapi import APIRouter, HTTPException, Request
from fastapi.responses import JSONResponse

from .audit import AuditLogger
from .models import PresentationRequestV27
from .idempotency import OperationBusy, PilotSerialGuard
from .providers import (
    AssetFolderNotFound,
    DriveProvider,
    OutputConflict,
    OutputFolderNotFound,
    ProviderError,
)
from .renderer import CapaRendererV27, RenderError, governed_filename


def _fingerprint(payload: PresentationRequestV27) -> str:
    encoded = json.dumps(
        payload.model_dump(mode="json"),
        ensure_ascii=False,
        sort_keys=True,
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def build_v27_router(
    audit_logger: AuditLogger,
    drive_provider: DriveProvider | None = None,
    renderer: CapaRendererV27 | None = None,
    serial_guard: PilotSerialGuard | None = None,
) -> APIRouter:
    router = APIRouter(prefix="/v2.7", tags=["V2.7"])
    serial_guard = serial_guard or PilotSerialGuard()

    @router.post("/gerar-apresentacao-estate-gover")
    def handle_v27_request(payload: PresentationRequestV27, request: Request):
        operation_id = request.state.operation_id
        base_event = {
            "operation_id": operation_id,
            "codigo_ativo": payload.codigo_ativo,
            "folder_id": payload.folder_id,
            "tipo_saida": payload.tipo_saida.value,
            "versao_saida": payload.versao_saida,
            "status_arquivo": "RASCUNHO",
            "publicacao_automatica": False,
        }
        audit_logger.record({
            **base_event,
            "event": "governed_request_validated",
            "result": "validado",
            "input_fingerprint": _fingerprint(payload),
        })

        # Fase 1 continua sendo o default seguro.
        if drive_provider is None or renderer is None:
            return JSONResponse(
                status_code=501,
                content={
                    "status": "nao_implementado",
                    "versao": "v2.7",
                    "fase": "FASE_1",
                    **base_event,
                    "arquivos": [],
                    "detail": (
                        "Contrato V2.7 validado. Providers de Fase 2 "
                        "não foram injetados."
                    ),
                },
            )

        if payload.tipo_saida.value != "CAPA":
            audit_logger.record({
                **base_event,
                "event": "phase2_scope_blocked",
                "result": "erro",
                "erro": "Fase 2 habilita somente CAPA",
            })
            raise HTTPException(
                status_code=501,
                detail="V2.7 Fase 2 habilita somente CAPA.",
            )

        filename = governed_filename(payload)

        claim = serial_guard.claim(
            f"{payload.folder_id}:{payload.tipo_saida.value}:{filename}"
        )
        try:
            claim.__enter__()
        except OperationBusy as exc:
            audit_logger.record({
                **base_event,
                "event": "pilot_serial_blocked",
                "result": "erro",
                "output_filename": filename,
            })
            raise HTTPException(status_code=409, detail=str(exc)) from exc

        try:
            drive_provider.validate_asset_folder(payload.folder_id)
            output_folder = drive_provider.resolve_output_folder(
                payload.folder_id, "CAPA"
            )

            if drive_provider.file_exists(payload.folder_id, filename):
                audit_logger.record({
                    **base_event,
                    "event": "destination_conflict",
                    "result": "erro",
                    "output_filename": filename,
                    "output_folder_id": output_folder,
                })
                raise HTTPException(
                    status_code=409,
                    detail="Arquivo já existe; sobrescrita não autorizada.",
                )

            audit_logger.record({
                **base_event,
                "event": "destination_validated",
                "result": "ok",
                "output_filename": filename,
                "output_folder_id": output_folder,
            })

            with tempfile.TemporaryDirectory(prefix="estate-gover-v27-") as tmp:
                local_path = Path(tmp) / filename
                render_result = renderer.render(payload, local_path)
                audit_logger.record({
                    **base_event,
                    "event": "render_completed",
                    "result": "ok",
                    "output_filename": filename,
                    "slide_count": render_result["slide_count"],
                    "hyperlinks_ok": render_result["hyperlinks_ok"],
                })

                # Second explicit conflict check immediately before upload.
                # GoogleDriveProvider repeats it internally as defense in depth.
                if drive_provider.file_exists(payload.folder_id, filename):
                    audit_logger.record({
                        **base_event,
                        "event": "pre_upload_conflict",
                        "result": "erro",
                        "output_filename": filename,
                        "output_folder_id": output_folder,
                    })
                    raise HTTPException(
                        status_code=409,
                        detail="Arquivo surgiu durante a geração; upload bloqueado.",
                    )

                audit_logger.record({
                    **base_event,
                    "event": "pre_upload_conflict_check",
                    "result": "ok",
                    "output_filename": filename,
                    "output_folder_id": output_folder,
                })

                uploaded = drive_provider.upload_draft(
                    payload.folder_id, local_path
                )

            audit_logger.record({
                **base_event,
                "event": "upload_completed",
                "result": "ok",
                "output_filename": filename,
                "output_folder_id": uploaded["parent_folder_id"],
                "drive_file_id": uploaded["file_id"],
            })

            return {
                "status": "rascunho_gerado",
                "versao": "v2.7",
                "fase": "FASE_2_HOMOLOGACAO",
                **base_event,
                "arquivos": [{
                    "produto": "CAPA",
                    "arquivo": uploaded["name"],
                    "drive_file_id": uploaded["file_id"],
                    "url": uploaded["webViewLink"],
                }],
            }

        except HTTPException:
            raise
        except (
            AssetFolderNotFound,
            OutputFolderNotFound,
            OutputConflict,
            ProviderError,
            RenderError,
        ) as exc:
            audit_logger.record({
                **base_event,
                "event": "phase2_failed",
                "result": "erro",
                "output_filename": filename,
                "erro": str(exc),
            })
            raise HTTPException(status_code=422, detail=str(exc)) from exc
        except Exception as exc:
            audit_logger.record({
                **base_event,
                "event": "phase2_failed",
                "result": "erro",
                "output_filename": filename,
                "error_type": type(exc).__name__,
            })
            raise HTTPException(
                status_code=500,
                detail="Falha interna sanitizada na operação V2.7.",
            ) from exc
        finally:
            claim.__exit__(None, None, None)

    return router
