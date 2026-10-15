"""
Batch mail-triage tool for mcp-microsoft.

role_obot#23. Ported from mpalermiti/outlook-mcp's tools/batch.py. The
mpalermiti original builds a raw ``kiota_abstractions.RequestInformation``
POST to ``/$batch`` and sends it through the SDK client's
``request_adapter`` directly (bypassing the typed builders, which have no
``$batch`` surface). This port drops that entirely in favor of
``GraphClient.batch()`` — guinacio's own ``$batch`` chunking helper already
used by ``tools/mail.py``'s ``bulk_move_emails``/``bulk_trash_emails`` — so
batch_triage gets the exact same auth, retry, and chunking behavior as the
rest of this codebase's bulk operations.

Implemented:
  - batch_triage
"""

from __future__ import annotations

from typing import Any, Literal
from urllib.parse import quote

from pydantic import model_validator

from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import WRITE_TOOL, register_tool
from mcp_microsoft.graph import get_graph
from mcp_microsoft.models import BatchTriageResponse, BatchTriageResultItem

_VALID_FLAG_STATUSES = ("flagged", "complete", "notFlagged")

# Graph's own $batch limit per HTTP call. GraphClient.batch() already chunks
# into groups of 20 for larger lists, but a single batch_triage call is kept
# at this ceiling (matching mpalermiti) so one tool call maps to one Graph
# round-trip rather than silently fanning out into several.
_MAX_BATCH_SIZE = 20


class BatchTriageInput(ToolRequestModel):
    message_ids: list[str]
    action: Literal["move", "flag", "categorize", "mark_read"]
    value: str
    profile: str | None = None

    @model_validator(mode="after")
    def _validate(self) -> "BatchTriageInput":
        if not self.message_ids:
            raise ValueError("message_ids must not be empty")
        if len(self.message_ids) > _MAX_BATCH_SIZE:
            raise ValueError(f"Maximum {_MAX_BATCH_SIZE} messages per batch (Graph API limit)")
        if self.action == "flag" and self.value not in _VALID_FLAG_STATUSES:
            raise ValueError(
                f"flag value must be one of {_VALID_FLAG_STATUSES}; got {self.value!r}"
            )
        return self


def _build_subrequest(index: int, message_id: str, action: str, value: Any) -> dict:
    """Map one (message_id, action, value) triage instruction to a $batch sub-request."""
    request_id = str(index)
    encoded_id = quote(message_id, safe="")

    if action == "move":
        return {
            "id": request_id,
            "method": "POST",
            "url": f"/me/messages/{encoded_id}/move",
            "body": {"destinationId": value},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "flag":
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{encoded_id}",
            "body": {"flag": {"flagStatus": value}},
            "headers": {"Content-Type": "application/json"},
        }
    if action == "categorize":
        categories = [c.strip() for c in value.split(",")] if isinstance(value, str) else value
        return {
            "id": request_id,
            "method": "PATCH",
            "url": f"/me/messages/{encoded_id}",
            "body": {"categories": categories},
            "headers": {"Content-Type": "application/json"},
        }
    # mark_read
    is_read = value.lower() == "true" if isinstance(value, str) else bool(value)
    return {
        "id": request_id,
        "method": "PATCH",
        "url": f"/me/messages/{encoded_id}",
        "body": {"isRead": is_read},
        "headers": {"Content-Type": "application/json"},
    }


async def batch_triage(params: BatchTriageInput) -> BatchTriageResponse:
    """
    Apply one triage action to up to 20 messages in a single Graph $batch call.

    10-20x faster than per-message calls for bulk triage — one HTTP round-trip
    to Graph instead of N.

    Args:
        message_ids: Up to 20 Graph message IDs.
        action: One of move, flag, categorize, mark_read.
        value: Action-specific — for "move" a destination folder ID (well-known
               name or Graph ID, passed straight through as destinationId, same
               as move_email); for "flag" one of flagged/complete/notFlagged;
               for "categorize" a comma-separated category list; for
               "mark_read" "true"/"false".
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)

    subrequests = [
        _build_subrequest(i, mid, params.action, params.value)
        for i, mid in enumerate(params.message_ids)
    ]

    responses = await g.batch(subrequests)

    id_to_message = {str(i): mid for i, mid in enumerate(params.message_ids)}
    results_by_id: dict[str, BatchTriageResultItem] = {}
    for resp in responses:
        mid = id_to_message.get(str(resp.get("id")))
        if mid is None:
            continue
        status = resp.get("status", 0)
        if 200 <= status < 300:
            results_by_id[mid] = BatchTriageResultItem(message_id=mid, status="success")
        else:
            body = resp.get("body") or {}
            err = body.get("error") or {}
            message = err.get("message") or f"HTTP {status}"
            results_by_id[mid] = BatchTriageResultItem(
                message_id=mid, status="error", error=message
            )

    results = [
        results_by_id.get(
            mid,
            BatchTriageResultItem(message_id=mid, status="error", error="No response from batch"),
        )
        for mid in params.message_ids
    ]
    succeeded = sum(1 for r in results if r.status == "success")

    return BatchTriageResponse(
        success=succeeded == len(results),
        action="batch_triage",
        action_applied=params.action,
        total=len(results),
        succeeded=succeeded,
        failed=len(results) - succeeded,
        results=results,
    )


def register(server) -> None:
    """Register the batch-triage tool with the given FastMCP server instance."""
    register_tool(server, batch_triage, annotations=WRITE_TOOL)
