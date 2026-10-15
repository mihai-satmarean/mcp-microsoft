"""
Delta-query (incremental sync) tools for mcp-microsoft.

role_obot#23. Ported from mpalermiti/outlook-mcp's tools/mail_delta.py,
tools/calendar_delta.py, and tools/contacts_delta.py. The mpalermiti originals
deliberately bypass the msgraph SDK and hit Graph's delta endpoints with raw
httpx + an ``azure-identity`` credential's bearer token — the SDK's typed
delta builders rebuild URL templates from query-parameter dataclasses (which
strips ``$deltatoken`` off a deltaLink) and discard the ``@removed``
annotation tombstones need.

This port keeps "raw-ish" Graph calls but drops the standalone httpx client
and bearer-minting: every page fetch instead goes through
``graph.get_graph(profile).get(path, headers=...)``, exactly like
``tools/mail.py``'s own ``_collect_folder_message_ids`` already does when it
follows a mail ``@odata.nextLink`` — extracting the ``/v1.0/...`` relative
path+query from the full URL Graph hands back, rather than requesting it
absolutely. ``GraphClient`` already carries the stdio/http identity, retry,
and auth-error handling every other tool gets, so delta calls get it too.

A caller-supplied ``delta_token`` is untrusted input (it round-trips through
whatever stored it between calls) that would otherwise let a tool be pointed
at an arbitrary absolute URL; ``_relative_path`` pins the host to
``graph.microsoft.com`` before it is used, mirroring mpalermiti's own
``require_graph_url`` guard in tools/_delta.py.

Per-endpoint Graph quirks, verified by mpalermiti's module docstrings and kept
here:
  - ``/me/mailFolders/{id}/messages/delta`` accepts ``$top``.
  - ``/me/calendarView/delta`` requires ``startDateTime``/``endDateTime`` on
    the first call (no whole-calendar sync) but rejects ``$top`` — page size
    goes in a ``Prefer: odata.maxpagesize=N`` header instead.
  - ``/me/contacts/delta`` accepts no query parameters at all on the first
    call; same ``Prefer`` header for page size.

Tokens are not persisted server-side — the caller stores the returned
``delta_token`` (a deltaLink once a sync round completes, or a nextLink mid-
round) and passes it back on the next call, the same pattern as the existing
pagination ``skip_token``/``next_page_token`` fields elsewhere in this
codebase.

Implemented:
  - list_emails_delta
  - list_events_delta
  - list_contacts_delta
"""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import quote, urlsplit

from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, register_tool
from mcp_microsoft.graph import get_graph
from mcp_microsoft.models import (
    ContactsDeltaResponse,
    DeltaItem,
    EventsDeltaResponse,
    MailDeltaResponse,
)

# Safety cap: stop auto-following nextLink pages within one tool call once
# this many items have accumulated, handing back the nextLink as delta_token
# so the caller resumes instead of one call draining an unbounded mailbox.
_PAGE_CAP_ITEMS = 500

_GRAPH_HOST = "graph.microsoft.com"
_V1_PATH_RE = re.compile(r"/v1\.0(/.+)$")


def _relative_path(url: str) -> str:
    """Return the ``/v1.0``-relative path+query of a Graph nextLink/deltaLink.

    Validates the host is Graph's own before trusting the path — a
    caller-supplied ``delta_token`` is untrusted input that would otherwise
    let this tool be pointed at an arbitrary URL via a shared, authenticated
    GraphClient. See the module docstring.
    """
    parsed = urlsplit(url)
    if parsed.scheme.lower() != "https" or parsed.netloc.lower() != _GRAPH_HOST:
        raise ValueError(
            "delta_token must be a graph.microsoft.com URL returned by a "
            "previous call to this tool; refusing to follow a foreign URL."
        )
    match = _V1_PATH_RE.search(url)
    if not match:
        raise ValueError("delta_token is not a recognizable Graph v1.0 URL.")
    return match.group(1)


def _tag_delta_item(raw: dict) -> dict:
    """Collapse a raw Graph delta item to the DeltaItem wire shape.

    Items annotated with ``@removed`` are tombstones — Graph returns only
    ``id`` (and the annotation), no other fields; those collapse to
    ``{id, is_deleted: True}``. Live items pass through with their raw Graph
    JSON fields intact (DeltaItem is a FlexibleMCPModel) plus
    ``is_deleted: False``.
    """
    if "@removed" in raw:
        return {"id": raw.get("id"), "is_deleted": True}
    return {**raw, "is_deleted": False}


async def _drain_delta(
    g: Any,
    *,
    initial_path: str | None,
    delta_token: str | None,
    headers: dict[str, str] | None = None,
) -> tuple[list[dict], str | None, bool]:
    """Walk Graph delta pages until a ``deltaLink`` or the safety cap.

    Returns ``(raw_items, next_token, has_more)``:
      - ``next_token`` is a deltaLink (sync round complete — store it and use
        it for the *next* round) or a nextLink (cap hit mid-round — pass it
        straight back in as delta_token to resume) or ``None`` (no changes
        and no new link; reuse the previous token).
      - ``has_more`` is True only on the cap-hit case.
    """
    path = _relative_path(delta_token) if delta_token else initial_path
    if path is None:
        raise ValueError(
            "Either delta_token or the first-call parameters (folder/window) are required."
        )

    collected: list[dict] = []
    while True:
        result = await g.get(path, headers=headers) or {}
        collected.extend(result.get("value") or [])

        delta_link = result.get("@odata.deltaLink")
        if delta_link:
            _relative_path(delta_link)  # validate shape before handing it back as a token
            return collected, delta_link, False

        next_link = result.get("@odata.nextLink")
        if not next_link:
            # No changes, no new link — valid (Graph may echo an unchanged
            # window). Caller keeps using their existing token.
            return collected, None, False

        if len(collected) >= _PAGE_CAP_ITEMS:
            return collected, next_link, True

        path = _relative_path(next_link)


# ---------------------------------------------------------------------------
# list_emails_delta
# ---------------------------------------------------------------------------


class ListEmailsDeltaInput(ToolRequestModel):
    folder: str = "inbox"
    page_size: int = 50
    delta_token: str | None = None
    profile: str | None = None


async def list_emails_delta(params: ListEmailsDeltaInput) -> MailDeltaResponse:
    """
    List mail changes in a folder since the last delta token (incremental sync).

    First call (delta_token omitted) returns a full snapshot of the folder.
    Pass the returned delta_token back on the next call to get only what
    changed since then. Tombstones (deleted messages) come back as
    ``{"id": ..., "is_deleted": true}`` with no other fields.

    Args:
        folder: Mail folder (well-known name or Graph ID). Ignored once
                delta_token is set — the token already encodes the folder.
        page_size: Items per Graph page (1-100), mapped to $top.
        delta_token: Opaque cursor from a previous call. Omit for the first call.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    page_size = max(1, min(params.page_size, 100))

    initial_path = None
    if not params.delta_token:
        initial_path = f"/me/mailFolders/{params.folder}/messages/delta?$top={page_size}"

    items, next_token, has_more = await _drain_delta(
        g, initial_path=initial_path, delta_token=params.delta_token
    )

    return MailDeltaResponse(
        folder=params.folder,
        count=len(items),
        messages=[DeltaItem.model_validate(_tag_delta_item(item)) for item in items],
        delta_token=next_token,
        has_more=has_more,
    )


# ---------------------------------------------------------------------------
# list_events_delta
# ---------------------------------------------------------------------------


class ListEventsDeltaInput(ToolRequestModel):
    start: str | None = None
    end: str | None = None
    page_size: int = 50
    delta_token: str | None = None
    profile: str | None = None


async def list_events_delta(params: ListEventsDeltaInput) -> EventsDeltaResponse:
    """
    List calendar changes within [start, end] since the last delta token.

    First call requires both start and end (ISO 8601) — Graph's
    /me/calendarView/delta has no whole-calendar sync, only a windowed one.
    Subsequent calls (delta_token set) ignore start/end; the token already
    encodes the window. Graph rejects $top on this endpoint — page_size is
    sent as a ``Prefer: odata.maxpagesize`` header instead.

    Args:
        start: ISO 8601 window start. Required on the first call.
        end: ISO 8601 window end. Required on the first call.
        page_size: Items per Graph page (1-100).
        delta_token: Opaque cursor from a previous call. Omit for the first call.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    page_size = max(1, min(params.page_size, 100))
    headers = {"Prefer": f"odata.maxpagesize={page_size}"}

    initial_path = None
    if not params.delta_token:
        if not params.start or not params.end:
            raise ValueError(
                "list_events_delta requires both 'start' and 'end' (ISO 8601) on "
                "the first call — Graph's calendarView delta has no whole-calendar sync."
            )
        initial_path = (
            f"/me/calendarView/delta?startDateTime={quote(params.start, safe='')}"
            f"&endDateTime={quote(params.end, safe='')}"
        )

    items, next_token, has_more = await _drain_delta(
        g, initial_path=initial_path, delta_token=params.delta_token, headers=headers
    )

    return EventsDeltaResponse(
        count=len(items),
        events=[DeltaItem.model_validate(_tag_delta_item(item)) for item in items],
        delta_token=next_token,
        has_more=has_more,
    )


# ---------------------------------------------------------------------------
# list_contacts_delta
# ---------------------------------------------------------------------------


class ListContactsDeltaInput(ToolRequestModel):
    page_size: int = 50
    delta_token: str | None = None
    profile: str | None = None


async def list_contacts_delta(params: ListContactsDeltaInput) -> ContactsDeltaResponse:
    """
    List contact changes since the last delta token.

    /me/contacts/delta accepts no query parameters at all on the first call;
    page size is sent as a ``Prefer: odata.maxpagesize`` header instead.

    Args:
        page_size: Items per Graph page (1-100).
        delta_token: Opaque cursor from a previous call. Omit for the first call.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    page_size = max(1, min(params.page_size, 100))
    headers = {"Prefer": f"odata.maxpagesize={page_size}"}

    initial_path = None if params.delta_token else "/me/contacts/delta"

    items, next_token, has_more = await _drain_delta(
        g, initial_path=initial_path, delta_token=params.delta_token, headers=headers
    )

    return ContactsDeltaResponse(
        count=len(items),
        contacts=[DeltaItem.model_validate(_tag_delta_item(item)) for item in items],
        delta_token=next_token,
        has_more=has_more,
    )


# ---------------------------------------------------------------------------
# Registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all delta-sync tools with the given FastMCP server instance."""
    register_tool(server, list_emails_delta, annotations=READ_ONLY_TOOL)
    register_tool(server, list_events_delta, annotations=READ_ONLY_TOOL)
    register_tool(server, list_contacts_delta, annotations=READ_ONLY_TOOL)
