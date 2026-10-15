"""
To Do (Microsoft To Do tasks) tools for mcp-microsoft.

role_obot#23. Ported from mpalermiti/outlook-mcp's tools/todo.py and
tools/todo_attachments.py (14 tools total). mpalermiti talks to Graph through
the Kiota-generated msgraph SDK (``graph_client.me.todo.lists...`` typed
request builders) under a raw ``azure-identity`` ``DeviceCodeCredential``.
Neither survives the port: this file uses guinacio's own REST seam instead —
``graph.get_graph(profile)`` / ``GraphClient.get/post/patch/delete/get_raw`` —
exactly like every other tool module in this codebase, so To Do tools get the
same stdio ``ProfileTokenProvider`` / http ``OboTokenProvider`` identity
handling, retry, and error masking as the mail or calendar tools.

Default-list resolution deliberately does NOT cache the resolved list id (the
mpalermiti original caches it in a ``WeakKeyDictionary`` keyed by the SDK
client). In guinacio's http (multi-user) transport, ``GraphClient`` instances
are shared across every signed-in user (see ``graph.get_graph``) — caching
"the default list" on that shared instance would leak one user's resolved
list id into another user's request. Re-resolving on every call costs one
extra Graph GET; it is the only safe behavior for a shared client.

Implemented:
  - list_task_lists
  - list_tasks
  - get_task
  - create_task
  - update_task
  - complete_task
  - delete_task
  - add_checklist_item
  - update_checklist_item
  - delete_checklist_item
  - list_task_attachments
  - download_task_attachment
  - upload_task_attachment
  - delete_task_attachment
"""

from __future__ import annotations

import base64
import mimetypes
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from mcp_microsoft.common.formatting import format_size_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import (
    DESTRUCTIVE_TOOL,
    READ_ONLY_TOOL,
    WRITE_TOOL,
    register_tool,
)
from mcp_microsoft.config import get_app_config
from mcp_microsoft.feature_flags import is_deletion_disabled
from mcp_microsoft.graph import get_graph
from mcp_microsoft.models import (
    AddChecklistItemResponse,
    ChecklistItemInfo,
    CompleteTodoTaskResponse,
    CreateTodoTaskResponse,
    DeleteChecklistItemResponse,
    DeleteTodoAttachmentResponse,
    DeleteTodoTaskResponse,
    DownloadTodoAttachmentResponse,
    ListTodoAttachmentsResponse,
    ListTodoListsResponse,
    ListTodoTasksResponse,
    TodoAttachmentInfo,
    TodoListInfo,
    TodoTaskDetailResponse,
    TodoTaskSummary,
    UpdateChecklistItemResponse,
    UpdateTodoTaskResponse,
    UploadTodoAttachmentResponse,
)

_VALID_STATUSES = {"notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"}
_VALID_IMPORTANCES = {"low", "normal", "high"}

# 20 MiB, not Graph's nominal 25 MB per taskFileAttachment: the inline POST is
# one JSON body, Graph rejects bodies over 30 MB, and base64 inflates the
# payload 4/3 — 25 MiB of file is ~33.4 MB of JSON and a guaranteed 400. Same
# derivation as mpalermiti's tools/todo_attachments.py module docstring.
_MAX_TASK_ATTACHMENT_BYTES = 20 * 1024 * 1024


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------


async def _resolve_list_id(g: Any, list_id: str | None) -> str:
    """Resolve *list_id* to a concrete Graph id, defaulting to the owner's default list.

    Single ``$top=100`` page, no caching (see module docstring). Falls back to
    the first list returned when no list is flagged as the wellknown default —
    matches mpalermiti's fallback behavior for mailboxes without one.
    """
    if list_id:
        return list_id

    result = await g.get("/me/todo/lists", params={"$top": 100}) or {}
    lists = result.get("value") or []
    first_id: str | None = None
    for lst in lists:
        if first_id is None:
            first_id = lst.get("id")
        if lst.get("isOwner") and lst.get("wellknownListName") == "defaultList":
            return lst["id"]

    if first_id is None:
        raise ValueError("No task lists found. Create a list in Microsoft To Do first.")
    return first_id


def _format_task(task: dict) -> TodoTaskSummary:
    due = task.get("dueDateTime") or {}
    completed = task.get("completedDateTime") or {}
    return TodoTaskSummary(
        id=task.get("id") or "",
        title=task.get("title") or "",
        status=task.get("status") or "notStarted",
        importance=task.get("importance") or "normal",
        due=due.get("dateTime"),
        completed=completed.get("dateTime"),
        created_at=task.get("createdDateTime") or "",
        is_reminder_on=bool(task.get("isReminderOn")),
        has_recurrence=task.get("recurrence") is not None,
    )


def _format_checklist_item(item: dict) -> ChecklistItemInfo:
    return ChecklistItemInfo(
        id=item.get("id") or "",
        display_name=item.get("displayName") or "",
        is_checked=bool(item.get("isChecked")),
        checked_at=item.get("checkedDateTime") or "",
        created_at=item.get("createdDateTime") or "",
    )


def _checked_last(items: list[ChecklistItemInfo]) -> list[ChecklistItemInfo]:
    """Unchecked items first, then by creation time — matches every To Do client surface."""
    return sorted(items, key=lambda i: (i.is_checked, i.created_at))


def _skiptoken_from_nextlink(next_link: str | None) -> str | None:
    """Extract ``$skiptoken`` from a Graph ``@odata.nextLink``, mirroring list_emails."""
    if not next_link:
        return None
    qs = parse_qs(urlparse(next_link).query)
    return qs.get("$skiptoken", [None])[0]


def _require_stdio(tool_name: str) -> None:
    """Reject attachment read/write-to-disk tools in http (multi-user) transport."""
    if get_app_config().transport == "http":
        raise ValueError(
            f"{tool_name} is not available in multi-user http mode (the server's "
            "disk is not the caller's disk); it is only registered in stdio mode."
        )


# ---------------------------------------------------------------------------
# list_task_lists
# ---------------------------------------------------------------------------


class ListTaskListsInput(ToolRequestModel):
    profile: str | None = None


async def list_task_lists(params: ListTaskListsInput) -> ListTodoListsResponse:
    """
    List all Microsoft To Do task lists for the signed-in user.

    Args:
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    result = await g.get("/me/todo/lists", params={"$top": 100}) or {}
    lists = result.get("value") or []

    task_lists = [
        TodoListInfo(
            id=lst.get("id") or "",
            display_name=lst.get("displayName") or "",
            is_default=bool(
                lst.get("isOwner") and lst.get("wellknownListName") == "defaultList"
            ),
        )
        for lst in lists
    ]
    return ListTodoListsResponse(count=len(task_lists), task_lists=task_lists)


# ---------------------------------------------------------------------------
# list_tasks
# ---------------------------------------------------------------------------


class ListTasksInput(ToolRequestModel):
    list_id: str | None = None
    status: str | None = None
    max_results: int = 25
    skip_token: str | None = None
    profile: str | None = None


async def list_tasks(params: ListTasksInput) -> ListTodoTasksResponse:
    """
    List tasks in a To Do list, newest first.

    Args:
        list_id: Target list ID. Omit for the signed-in user's default list.
        status: Optional filter — one of notStarted, inProgress, completed,
                waitingOnOthers, deferred.
        max_results: Items per page (1-100).
        skip_token: Opaque pagination cursor returned as next_page_token from a
                    previous call. Omit for the first page.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)

    query: dict[str, Any] = {
        "$top": max(1, min(params.max_results, 100)),
        "$orderby": "createdDateTime desc",
    }
    if params.status:
        if params.status not in _VALID_STATUSES:
            raise ValueError(
                f"Invalid status {params.status!r}. Must be one of: {sorted(_VALID_STATUSES)}"
            )
        query["$filter"] = f"status eq '{params.status}'"
    if params.skip_token:
        query["$skiptoken"] = params.skip_token

    result = await g.get(f"/me/todo/lists/{resolved_id}/tasks", params=query) or {}
    tasks = [_format_task(t) for t in (result.get("value") or [])]
    next_page_token = _skiptoken_from_nextlink(result.get("@odata.nextLink"))

    return ListTodoTasksResponse(
        list_id=resolved_id,
        count=len(tasks),
        tasks=tasks,
        next_page_token=next_page_token,
        has_more=next_page_token is not None,
    )


# ---------------------------------------------------------------------------
# get_task
# ---------------------------------------------------------------------------


class GetTaskInput(ToolRequestModel):
    task_id: str
    list_id: str | None = None
    profile: str | None = None


async def get_task(params: GetTaskInput) -> TodoTaskDetailResponse:
    """
    Get full task details, including notes (body) and checklist items.

    Args:
        task_id: The Graph task ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)

    task = (
        await g.get(
            f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}",
            params={"$expand": "checklistItems"},
        )
        or {}
    )

    summary = _format_task(task)
    body = (task.get("body") or {}).get("content") or ""
    raw_items = task.get("checklistItems") or []
    items = _checked_last([_format_checklist_item(i) for i in raw_items])

    return TodoTaskDetailResponse(
        id=summary.id,
        list_id=resolved_id,
        title=summary.title,
        status=summary.status,
        importance=summary.importance,
        due=summary.due,
        completed=summary.completed,
        created_at=summary.created_at,
        is_reminder_on=summary.is_reminder_on,
        body=body,
        has_recurrence=summary.has_recurrence,
        checklist_items=items,
        checklist_count=len(items),
    )


# ---------------------------------------------------------------------------
# create_task
# ---------------------------------------------------------------------------


class CreateTaskInput(ToolRequestModel):
    title: str
    list_id: str | None = None
    due: str | None = None
    importance: str | None = None
    body: str | None = None
    reminder: bool | None = None
    profile: str | None = None


async def create_task(params: CreateTaskInput) -> CreateTodoTaskResponse:
    """
    Create a task in a To Do list.

    Args:
        title: Task title.
        list_id: Target list ID. Omit for the signed-in user's default list.
        due: Optional ISO 8601 due datetime (UTC).
        importance: Optional — one of low, normal, high.
        body: Optional plain-text note.
        reminder: Optional reminder flag. Requires `due` — Graph silently stores
                  isReminderOn=false unless reminderDateTime is also set, so
                  there is nothing to anchor the reminder on without a due date.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    if params.importance is not None and params.importance not in _VALID_IMPORTANCES:
        raise ValueError(
            f"Invalid importance {params.importance!r}. Must be one of: {sorted(_VALID_IMPORTANCES)}"
        )
    if params.reminder and not params.due:
        raise ValueError("reminder=True needs a `due` datetime to anchor the reminder on.")

    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)

    payload: dict[str, Any] = {"title": params.title}
    if params.due:
        payload["dueDateTime"] = {"dateTime": params.due, "timeZone": "UTC"}
    if params.importance:
        payload["importance"] = params.importance
    if params.body:
        payload["body"] = {"content": params.body, "contentType": "text"}
    if params.reminder is not None:
        payload["isReminderOn"] = params.reminder
        if params.reminder:
            payload["reminderDateTime"] = {"dateTime": params.due, "timeZone": "UTC"}

    result = await g.post(f"/me/todo/lists/{resolved_id}/tasks", json=payload) or {}
    if not result.get("id"):
        # The task may exist server-side even on an empty 201/204 response; an
        # agent's natural retry here would create a duplicate.
        raise ValueError(
            "Task was created but Graph returned no id for it — do not retry "
            "blindly (that would create a duplicate); find it with list_tasks instead."
        )

    return CreateTodoTaskResponse(
        success=True,
        action="create_task",
        task_id=result["id"],
        list_id=resolved_id,
        title=result.get("title") or params.title,
    )


# ---------------------------------------------------------------------------
# update_task
# ---------------------------------------------------------------------------


class UpdateTaskInput(ToolRequestModel):
    task_id: str
    list_id: str | None = None
    title: str | None = None
    due: str | None = None
    body: str | None = None
    importance: str | None = None
    profile: str | None = None


async def update_task(params: UpdateTaskInput) -> UpdateTodoTaskResponse:
    """
    Update a task's fields. Only the fields you pass change.

    Args:
        task_id: The Graph task ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        title / due / body / importance: Fields to update; omit to leave unchanged.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    if params.importance is not None and params.importance not in _VALID_IMPORTANCES:
        raise ValueError(
            f"Invalid importance {params.importance!r}. Must be one of: {sorted(_VALID_IMPORTANCES)}"
        )

    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)

    payload: dict[str, Any] = {}
    updated_fields: list[str] = []
    if params.title is not None:
        payload["title"] = params.title
        updated_fields.append("title")
    if params.due is not None:
        payload["dueDateTime"] = {"dateTime": params.due, "timeZone": "UTC"}
        updated_fields.append("due")
    if params.body is not None:
        payload["body"] = {"content": params.body, "contentType": "text"}
        updated_fields.append("body")
    if params.importance is not None:
        payload["importance"] = params.importance
        updated_fields.append("importance")

    if not payload:
        raise ValueError("Provide at least one field to update.")

    await g.patch(f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}", json=payload)

    return UpdateTodoTaskResponse(
        success=True,
        action="update_task",
        task_id=params.task_id,
        list_id=resolved_id,
        updated_fields=updated_fields,
        updated_fields_display=", ".join(updated_fields),
    )


# ---------------------------------------------------------------------------
# complete_task
# ---------------------------------------------------------------------------


class CompleteTaskInput(ToolRequestModel):
    task_id: str
    list_id: str | None = None
    profile: str | None = None


async def complete_task(params: CompleteTaskInput) -> CompleteTodoTaskResponse:
    """
    Mark a task as completed.

    Args:
        task_id: The Graph task ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)

    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.0000000Z")
    payload = {
        "status": "completed",
        "completedDateTime": {"dateTime": now_iso, "timeZone": "UTC"},
    }
    await g.patch(f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}", json=payload)

    return CompleteTodoTaskResponse(
        success=True,
        action="complete_task",
        task_id=params.task_id,
        list_id=resolved_id,
    )


# ---------------------------------------------------------------------------
# delete_task
# ---------------------------------------------------------------------------


class DeleteTaskInput(ToolRequestModel):
    task_id: str
    list_id: str | None = None
    profile: str | None = None


async def delete_task(params: DeleteTaskInput) -> DeleteTodoTaskResponse:
    """
    Delete a task from a To Do list. Irreversible.

    Args:
        task_id: The Graph task ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    await g.delete(f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}")
    return DeleteTodoTaskResponse(
        success=True, action="delete_task", task_id=params.task_id, list_id=resolved_id
    )


# ---------------------------------------------------------------------------
# add_checklist_item / update_checklist_item / delete_checklist_item
# ---------------------------------------------------------------------------


class AddChecklistItemInput(ToolRequestModel):
    task_id: str
    display_name: str
    list_id: str | None = None
    profile: str | None = None


async def add_checklist_item(params: AddChecklistItemInput) -> AddChecklistItemResponse:
    """
    Add a checklist item (sub-step) to a task.

    Args:
        task_id: The Graph task ID.
        display_name: Checklist item label (non-empty).
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    name = (params.display_name or "").strip()
    if not name:
        raise ValueError("display_name must be a non-empty string.")

    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)

    result = (
        await g.post(
            f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/checklistItems",
            json={"displayName": name},
        )
        or {}
    )
    if not result.get("id"):
        raise ValueError(
            "Checklist item was added but Graph returned no id for it — do not "
            "retry blindly (that would create a duplicate); read the task back "
            "with get_task instead."
        )

    return AddChecklistItemResponse(
        success=True,
        action="add_checklist_item",
        task_id=params.task_id,
        checklist_item_id=result["id"],
        display_name=result.get("displayName") or name,
    )


class UpdateChecklistItemInput(ToolRequestModel):
    task_id: str
    checklist_item_id: str
    display_name: str | None = None
    is_checked: bool | None = None
    list_id: str | None = None
    profile: str | None = None


async def update_checklist_item(
    params: UpdateChecklistItemInput,
) -> UpdateChecklistItemResponse:
    """
    Update a checklist item (partial patch — only the fields you pass change).

    ``checkedDateTime`` is maintained server-side from ``is_checked``; it is not
    settable directly here.

    Args:
        task_id: The Graph task ID.
        checklist_item_id: The Graph checklist item ID.
        display_name: New label, if renaming.
        is_checked: New checked state, if toggling.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    if params.display_name is None and params.is_checked is None:
        raise ValueError("Provide at least one of display_name or is_checked.")

    payload: dict[str, Any] = {}
    updated_fields: list[str] = []
    if params.display_name is not None:
        name = params.display_name.strip()
        if not name:
            raise ValueError("display_name must be a non-empty string.")
        payload["displayName"] = name
        updated_fields.append("display_name")
    if params.is_checked is not None:
        payload["isChecked"] = params.is_checked
        updated_fields.append("is_checked")

    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    await g.patch(
        f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/checklistItems/"
        f"{params.checklist_item_id}",
        json=payload,
    )

    return UpdateChecklistItemResponse(
        success=True,
        action="update_checklist_item",
        task_id=params.task_id,
        checklist_item_id=params.checklist_item_id,
        updated_fields=updated_fields,
    )


class DeleteChecklistItemInput(ToolRequestModel):
    task_id: str
    checklist_item_id: str
    list_id: str | None = None
    profile: str | None = None


async def delete_checklist_item(
    params: DeleteChecklistItemInput,
) -> DeleteChecklistItemResponse:
    """
    Delete a checklist item from a task.

    Args:
        task_id: The Graph task ID.
        checklist_item_id: The Graph checklist item ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    await g.delete(
        f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/checklistItems/"
        f"{params.checklist_item_id}"
    )
    return DeleteChecklistItemResponse(
        success=True,
        action="delete_checklist_item",
        task_id=params.task_id,
        checklist_item_id=params.checklist_item_id,
    )


# ---------------------------------------------------------------------------
# list_task_attachments / download_task_attachment / upload_task_attachment /
# delete_task_attachment
# ---------------------------------------------------------------------------


class ListTaskAttachmentsInput(ToolRequestModel):
    task_id: str
    list_id: str | None = None
    profile: str | None = None


async def list_task_attachments(
    params: ListTaskAttachmentsInput,
) -> ListTodoAttachmentsResponse:
    """
    List attachments on a To Do task.

    Args:
        task_id: The Graph task ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    result = (
        await g.get(
            f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/attachments",
            params={"$select": "id,name,size,contentType,lastModifiedDateTime"},
        )
        or {}
    )
    attachments = [
        TodoAttachmentInfo(
            id=att.get("id") or "",
            name=att.get("name") or "",
            size_bytes=att.get("size") or 0,
            size_display=format_size_display(att.get("size")),
            content_type=att.get("contentType") or "",
            last_modified_at=att.get("lastModifiedDateTime") or "",
        )
        for att in (result.get("value") or [])
    ]
    return ListTodoAttachmentsResponse(
        task_id=params.task_id, count=len(attachments), attachments=attachments
    )


class DownloadTaskAttachmentInput(ToolRequestModel):
    task_id: str
    attachment_id: str
    file_path: Path
    list_id: str | None = None
    profile: str | None = None


async def download_task_attachment(
    params: DownloadTaskAttachmentInput,
) -> DownloadTodoAttachmentResponse:
    """
    Download a To Do task attachment and write it to local disk.

    Mirrors ``attachments.save_attachment_to_path`` for mail: binary bytes come
    from Graph's ``$value`` endpoint, written straight to ``file_path`` with no
    base64 round-trip through the tool response. Not available in multi-user
    http mode.

    Args:
        task_id: The Graph task ID.
        attachment_id: The Graph attachment ID.
        file_path: Exact target file path, or an existing directory (uses the
                   attachment's own filename).
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    _require_stdio("download_task_attachment")

    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    base = f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/attachments/{params.attachment_id}"

    meta = await g.get(base, params={"$select": "id,name,size,contentType"}) or {}
    raw_bytes = await g.get_raw(f"{base}/$value")

    att_name = meta.get("name") or "attachment"
    content_type = meta.get("contentType") or "application/octet-stream"

    target = params.file_path
    if target.exists() and target.is_dir():
        safe_name = Path(att_name).name or "attachment"
        target = target / safe_name
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(raw_bytes)

    return DownloadTodoAttachmentResponse(
        success=True,
        action="download_task_attachment",
        task_id=params.task_id,
        attachment_id=params.attachment_id,
        file_path=str(target),
        filename=att_name,
        size_bytes=len(raw_bytes),
        size_display=format_size_display(len(raw_bytes)),
        content_type=content_type,
    )


class UploadTaskAttachmentInput(ToolRequestModel):
    task_id: str
    file_path: Path
    list_id: str | None = None
    profile: str | None = None


async def upload_task_attachment(
    params: UploadTaskAttachmentInput,
) -> UploadTodoAttachmentResponse:
    """
    Attach a local file to a To Do task via an inline base64 POST.

    Accepted size: 1 byte - 20 MiB (Graph caps $batch-free request bodies at
    30 MB and base64 inflates content 4/3). Not available in multi-user http
    mode.

    Args:
        task_id: The Graph task ID.
        file_path: Local file to attach.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    _require_stdio("upload_task_attachment")

    if not params.file_path.is_file():
        raise ValueError(f"Attachment file not found: {params.file_path}")

    content = params.file_path.read_bytes()
    if not content:
        raise ValueError("Attachment file is empty — accepted size is 1 byte - 20 MiB.")
    if len(content) > _MAX_TASK_ATTACHMENT_BYTES:
        raise ValueError(
            f"Attachment is over the 20 MiB ceiling ({len(content)} bytes read). Graph "
            "caps request bodies at 30 MB and base64 inflates the file 4/3."
        )

    content_type, _ = mimetypes.guess_type(str(params.file_path))

    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    payload = {
        "@odata.type": "#microsoft.graph.taskFileAttachment",
        "name": params.file_path.name,
        "contentType": content_type or "application/octet-stream",
        "size": len(content),
        "contentBytes": base64.b64encode(content).decode(),
    }
    result = (
        await g.post(f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/attachments", json=payload)
        or {}
    )
    if not result.get("id"):
        raise ValueError(
            "Attachment upload got no id back from Graph — do not retry blindly "
            "(it may have attached anyway); verify with list_task_attachments."
        )

    return UploadTodoAttachmentResponse(
        success=True,
        action="upload_task_attachment",
        task_id=params.task_id,
        attachment_id=result["id"],
        filename=result.get("name") or params.file_path.name,
        size_bytes=len(content),
    )


class DeleteTaskAttachmentInput(ToolRequestModel):
    task_id: str
    attachment_id: str
    list_id: str | None = None
    profile: str | None = None


async def delete_task_attachment(
    params: DeleteTaskAttachmentInput,
) -> DeleteTodoAttachmentResponse:
    """
    Remove an attachment from a To Do task.

    Args:
        task_id: The Graph task ID.
        attachment_id: The Graph attachment ID.
        list_id: Target list ID. Omit for the signed-in user's default list.
        profile: Microsoft 365 profile to use. Omit to use the default profile.
    """
    g = get_graph(params.profile)
    resolved_id = await _resolve_list_id(g, params.list_id)
    await g.delete(
        f"/me/todo/lists/{resolved_id}/tasks/{params.task_id}/attachments/{params.attachment_id}"
    )
    return DeleteTodoAttachmentResponse(
        success=True,
        action="delete_task_attachment",
        task_id=params.task_id,
        attachment_id=params.attachment_id,
    )


# ---------------------------------------------------------------------------
# Registration
# ---------------------------------------------------------------------------


def register(server) -> None:
    """Register all To Do tools with the given FastMCP server instance."""
    register_tool(server, list_task_lists, annotations=READ_ONLY_TOOL)
    register_tool(server, list_tasks, annotations=READ_ONLY_TOOL)
    register_tool(server, get_task, annotations=READ_ONLY_TOOL)
    register_tool(server, create_task, annotations=WRITE_TOOL)
    register_tool(server, update_task, annotations=WRITE_TOOL)
    register_tool(server, complete_task, annotations=WRITE_TOOL)
    register_tool(server, add_checklist_item, annotations=WRITE_TOOL)
    register_tool(server, update_checklist_item, annotations=WRITE_TOOL)
    register_tool(server, list_task_attachments, annotations=READ_ONLY_TOOL)

    if not is_deletion_disabled():
        register_tool(server, delete_task, annotations=DESTRUCTIVE_TOOL)
        register_tool(server, delete_checklist_item, annotations=DESTRUCTIVE_TOOL)
        register_tool(server, delete_task_attachment, annotations=DESTRUCTIVE_TOOL)

    if get_app_config().transport != "http":
        register_tool(server, download_task_attachment, annotations=WRITE_TOOL)
        register_tool(server, upload_task_attachment, annotations=WRITE_TOOL)
