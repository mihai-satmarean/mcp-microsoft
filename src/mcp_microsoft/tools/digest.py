"""
Composed "since last call" digest tool for mcp-microsoft.

role_obot#23. Ported from mpalermiti/outlook-mcp's tools/digest.py. Wraps
``tools/delta.py``'s three delta tools (mail / events / contacts) into a
single structured payload for a recurring agent loop (a morning brief or
hourly inbox sweep) that wants one digest call instead of orchestrating three
delta calls and reasoning over the raw item shapes itself.

Simplifications versus the mpalermiti original (documented, not silent):
  - No ``syncStateNotFound`` (Graph 410) auto-resync. A stale/expired
    delta_token surfaces as a plain httpx.HTTPStatusError from the underlying
    delta call; the caller should drop that resource's token and retry with
    it omitted (which re-bootstraps). mpalermiti recovers this automatically
    per-resource; porting that resync-with-partial-failure-isolation behavior
    faithfully would have roughly doubled this file for a tertiary feature,
    so it is left as a documented follow-up rather than ported partially.
  - The three resources are drained sequentially, not concurrently
    (``asyncio.gather`` in the original) — simpler control flow; a resource
    with a very large bootstrap snapshot delays the others by the same
    wall-clock amount either way since Graph calls dominate.

Implemented:
  - changes_since
"""

from __future__ import annotations

from datetime import datetime, timedelta, timezone

from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, register_tool
from mcp_microsoft.models import (
    ChangesSinceResponse,
    DeltaTokens,
    DigestContactsSection,
    DigestEventItem,
    DigestEventsSection,
    DigestMailSection,
    DigestUrgentMail,
)
from mcp_microsoft.tools.delta import (
    ListContactsDeltaInput,
    ListEmailsDeltaInput,
    ListEventsDeltaInput,
    list_contacts_delta,
    list_emails_delta,
    list_events_delta,
)

# Safety cap on internal pagination drain per resource. Each underlying delta
# call is already bounded to delta.py's _PAGE_CAP_ITEMS (500); this lets the
# digest chain up to 5 of those pages for a hard ceiling of ~2,500 items per
# resource per call — bounds a very large bootstrap snapshot.
_MAX_INTERNAL_PAGES = 5
_PAGE_SIZE = 50
_CALENDAR_LOOKAHEAD_DAYS = 7
_TOP_SENDERS = 5


class ChangesSinceInput(ToolRequestModel):
    mail_delta_token: str | None = None
    events_delta_token: str | None = None
    contacts_delta_token: str | None = None
    fallback_window_hours: int = 24
    folder: str = "inbox"
    profile: str | None = None


def _parse_iso(value: str) -> datetime | None:
    """Parse a Graph ISO 8601 string to an aware UTC datetime, or None on failure."""
    if not value:
        return None
    s = value.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


# ---------------------------------------------------------------------------
# Per-resource drain (follow has_more across delta.py's own page cap)
# ---------------------------------------------------------------------------


async def _drain_mail(
    token: str | None, folder: str, profile: str | None
) -> tuple[list[dict], str | None]:
    items: list[dict] = []
    cur_token = token
    last_token = token
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_emails_delta(
            ListEmailsDeltaInput(
                folder=folder, page_size=_PAGE_SIZE, delta_token=cur_token, profile=profile
            )
        )
        items.extend(item.model_dump() for item in page.messages)
        last_token = page.delta_token if page.delta_token is not None else last_token
        if not page.has_more:
            break
        cur_token = page.delta_token
    return items, last_token


async def _drain_events(
    token: str | None, bootstrap_start: str, bootstrap_end: str, profile: str | None
) -> tuple[list[dict], str | None]:
    items: list[dict] = []
    cur_token = token
    last_token = token
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_events_delta(
            ListEventsDeltaInput(
                start=None if cur_token else bootstrap_start,
                end=None if cur_token else bootstrap_end,
                page_size=_PAGE_SIZE,
                delta_token=cur_token,
                profile=profile,
            )
        )
        items.extend(item.model_dump() for item in page.events)
        last_token = page.delta_token if page.delta_token is not None else last_token
        if not page.has_more:
            break
        cur_token = page.delta_token
    return items, last_token


async def _drain_contacts(token: str | None, profile: str | None) -> tuple[list[dict], str | None]:
    items: list[dict] = []
    cur_token = token
    last_token = token
    for _ in range(_MAX_INTERNAL_PAGES):
        page = await list_contacts_delta(
            ListContactsDeltaInput(page_size=_PAGE_SIZE, delta_token=cur_token, profile=profile)
        )
        items.extend(item.model_dump() for item in page.contacts)
        last_token = page.delta_token if page.delta_token is not None else last_token
        if not page.has_more:
            break
        cur_token = page.delta_token
    return items, last_token


# ---------------------------------------------------------------------------
# Classifiers
# ---------------------------------------------------------------------------


def _classify_mail(items: list[dict], *, window_start: datetime | None) -> DigestMailSection:
    """Build the mail section. ``window_start`` filters a bootstrap snapshot
    (no token on the call that produced *items*) so a first call doesn't
    flood the digest with mailbox history; it is None on a token-driven call,
    where every returned item is already "since the last token" by definition.
    """
    new_count = 0
    removed_count = 0
    urgent: list[DigestUrgentMail] = []
    sender_counter: dict[str, int] = {}

    for item in items:
        if item.get("is_deleted"):
            removed_count += 1
            continue

        if window_start is not None:
            received_dt = _parse_iso(item.get("receivedDateTime") or "")
            if received_dt is None or received_dt < window_start:
                continue

        new_count += 1

        importance = (item.get("importance") or "").lower()
        flag_status = (item.get("flag") or {}).get("flagStatus") or ""
        is_high = importance == "high"
        is_flagged = flag_status == "flagged"
        from_address = ((item.get("from") or {}).get("emailAddress") or {}).get("address") or ""
        if is_high or is_flagged:
            urgent.append(
                DigestUrgentMail(
                    id=item.get("id") or "",
                    subject=item.get("subject") or "",
                    from_address=from_address,
                    received_at=item.get("receivedDateTime") or "",
                    is_high_importance=is_high,
                    is_flagged=is_flagged,
                )
            )
        if from_address:
            key = from_address.lower()
            sender_counter[key] = sender_counter.get(key, 0) + 1

    by_sender = dict(
        sorted(sender_counter.items(), key=lambda kv: kv[1], reverse=True)[:_TOP_SENDERS]
    )

    return DigestMailSection(
        new_count=new_count, removed_count=removed_count, urgent=urgent, by_sender=by_sender
    )


def _classify_events(items: list[dict]) -> DigestEventsSection:
    """Build the events section. A delta response can't distinguish new vs.
    modified on its own (no @changed annotation); everything live lands in
    "new", matching mpalermiti's documented behavior. Tombstones land in
    "cancelled".
    """
    new_list: list[DigestEventItem] = []
    cancelled_list: list[DigestEventItem] = []

    for item in items:
        if item.get("is_deleted"):
            cancelled_list.append(DigestEventItem(id=item.get("id") or ""))
            continue
        start_node = item.get("start") or {}
        end_node = item.get("end") or {}
        new_list.append(
            DigestEventItem(
                id=item.get("id") or "",
                subject=item.get("subject") or "",
                start=start_node.get("dateTime") or "",
                end=end_node.get("dateTime") or "",
            )
        )

    return DigestEventsSection(new=new_list, cancelled=cancelled_list)


def _classify_contacts(items: list[dict]) -> DigestContactsSection:
    new_count = sum(1 for item in items if not item.get("is_deleted"))
    removed_count = sum(1 for item in items if item.get("is_deleted"))
    return DigestContactsSection(new_count=new_count, removed_count=removed_count)


# ---------------------------------------------------------------------------
# changes_since
# ---------------------------------------------------------------------------


async def changes_since(params: ChangesSinceInput) -> ChangesSinceResponse:
    """
    Composed "since last call" digest across mail, calendar, and contacts.

    First call per resource (its *_delta_token omitted) bootstraps from a full
    Graph snapshot: mail is filtered to fallback_window_hours so the digest
    isn't flooded with history, and calendar bootstraps a
    [now - fallback_window_hours, now + 7 days] window. Store the returned
    delta_tokens and pass them back on the next call for a true incremental
    digest. Each resource's token is independent — a missing token for one
    resource doesn't block the other two.

    Args:
        mail_delta_token / events_delta_token / contacts_delta_token: Opaque
            cursors from a previous call's delta_tokens. Omit all three for a
            first/bootstrap call.
        fallback_window_hours: Bootstrap look-back window for mail (default 24h).
        folder: Mail folder to track (default inbox).
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    now = datetime.now(timezone.utc)
    window_start = now - timedelta(hours=max(0, params.fallback_window_hours))
    bootstrap_end_dt = now + timedelta(days=_CALENDAR_LOOKAHEAD_DAYS)
    bootstrap_start_iso = window_start.strftime("%Y-%m-%dT%H:%M:%SZ")
    bootstrap_end_iso = bootstrap_end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    mail_items, mail_token = await _drain_mail(
        params.mail_delta_token, params.folder, params.profile
    )
    events_items, events_token = await _drain_events(
        params.events_delta_token, bootstrap_start_iso, bootstrap_end_iso, params.profile
    )
    contacts_items, contacts_token = await _drain_contacts(
        params.contacts_delta_token, params.profile
    )

    mail_section = _classify_mail(
        mail_items, window_start=(window_start if params.mail_delta_token is None else None)
    )
    events_section = _classify_events(events_items)
    contacts_section = _classify_contacts(contacts_items)

    return ChangesSinceResponse(
        mail=mail_section,
        events=events_section,
        contacts=contacts_section,
        delta_tokens=DeltaTokens(mail=mail_token, events=events_token, contacts=contacts_token),
        window_start=bootstrap_start_iso,
        window_end=now.strftime("%Y-%m-%dT%H:%M:%SZ"),
    )


def register(server) -> None:
    """Register the composed digest tool with the given FastMCP server instance."""
    register_tool(server, changes_since, annotations=READ_ONLY_TOOL)
