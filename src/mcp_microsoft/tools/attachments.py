"""
Attachment tools for mcp-microsoft.

All tools use the Microsoft Graph API via the async graph client.

Implemented:
  - list_attachments
  - download_attachment
  - save_attachment_to_path (role_obot#23, ported from OraCool/outlook-mcp's
    tools/email_reader.py — see that function's docstring for the original
    behavior this mirrors)
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import Optional

from fastmcp.utilities.types import File

from mcp_microsoft.common.formatting import format_size_display
from mcp_microsoft.common.request_model import ToolRequestModel
from mcp_microsoft.common.tooling import READ_ONLY_TOOL, WRITE_TOOL, register_tool
from mcp_microsoft.config import get_app_config
from mcp_microsoft.graph_types import GraphAttachment, parse_graph_collection
from mcp_microsoft.models import (
    AttachmentInfo,
    DownloadAttachmentResponse,
    ListAttachmentsResponse,
    SaveAttachmentToPathResponse,
)
from mcp_microsoft.graph import get_graph

# ---------------------------------------------------------------------------
# list_attachments
# ---------------------------------------------------------------------------

class ListAttachmentsInput(ToolRequestModel):
    message_id: str
    profile: str | None = None


class DownloadAttachmentInput(ToolRequestModel):
    message_id: str
    attachment_id: str
    save_path: Path | None = None
    profile: str | None = None


class DownloadAttachmentHttpInput(ToolRequestModel):
    """HTTP input excludes server-local paths that remote callers cannot use."""

    message_id: str
    attachment_id: str
    profile: str | None = None


async def list_attachments(params: ListAttachmentsInput) -> ListAttachmentsResponse:
    """
    List all attachments on an email message.

    Args:
        message_id: The Graph message ID.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured attachment metadata.
    """
    g = get_graph(params.profile)
    query = {
        "$select": "id,name,size,contentType,isInline",
    }

    result = await g.get(f"/me/messages/{params.message_id}/attachments", params=query)
    attachments = parse_graph_collection(result, GraphAttachment)

    items: list[AttachmentInfo] = []
    for att in attachments:
        size_bytes = att.size
        items.append(
            AttachmentInfo(
                id=att.id,
                name=att.name,
                size_bytes=size_bytes,
                size_display=format_size_display(size_bytes),
                content_type=att.content_type or "unknown",
                is_inline=att.is_inline,
            )
        )

    return ListAttachmentsResponse(
        message_id=params.message_id,
        count=len(items),
        attachments=items,
    )


# ---------------------------------------------------------------------------
# download_attachment
# ---------------------------------------------------------------------------


async def download_attachment(
    params: DownloadAttachmentInput,
) -> DownloadAttachmentResponse | File:
    """
    Download an email attachment, saving it to disk or returning a file payload.

    When running in Claude Desktop, always provide save_path (e.g. the user's
    Downloads folder) so the file is written to disk. Omitting save_path returns
    a FastMCP file object, which is only useful for programmatic embedding.

    Args:
        message_id: The Graph message ID.
        attachment_id: The Graph attachment ID.
        save_path: Optional filesystem path to save the file. Omit to return a FastMCP file payload.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        A FastMCP file when no save path is given, or structured file-save metadata.
    """
    if params.save_path is not None and get_app_config().transport == "http":
        raise ValueError(
            "save_path is not available in multi-user http mode (the server's "
            "disk is not the caller's disk); omit it to receive the file inline."
        )

    g = get_graph(params.profile)
    result = GraphAttachment.model_validate(await g.get(
        f"/me/messages/{params.message_id}/attachments/{params.attachment_id}"
    ) or {})

    att_name = result.name or "attachment"
    content_type = result.content_type or "application/octet-stream"
    content_bytes_b64: Optional[str] = result.content_bytes

    if content_bytes_b64 is None:
        return DownloadAttachmentResponse(
            success=False,
            action="download_attachment",
            message_id=params.message_id,
            attachment_id=params.attachment_id,
            filename=att_name,
            error="Attachment has no downloadable content.",
        )

    raw_bytes = base64.b64decode(content_bytes_b64)

    if params.save_path is None:
        file_format = Path(att_name).suffix.lstrip(".")
        if not file_format and "/" in content_type:
            file_format = content_type.split("/", 1)[1]
        return File(data=raw_bytes, format=file_format or None, name=att_name)

    # Resolve final output path — sanitize remote filename to prevent traversal
    resolved_path = params.save_path
    if params.save_path.is_dir():
        safe_name = Path(att_name).name  # strip directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "attachment"
        resolved_path = params.save_path / safe_name

    resolved_path.parent.mkdir(parents=True, exist_ok=True)
    with resolved_path.open("wb") as fh:
        fh.write(raw_bytes)

    size_str = format_size_display(len(raw_bytes))
    return DownloadAttachmentResponse(
        success=True,
        action="download_attachment",
        message_id=params.message_id,
        attachment_id=params.attachment_id,
        path=str(resolved_path),
        filename=att_name,
        size_bytes=len(raw_bytes),
        size_display=size_str,
        content_type=content_type,
    )


async def download_attachment_http(
    params: DownloadAttachmentHttpInput,
) -> DownloadAttachmentResponse | File:
    """Download an email attachment and return it inline to the remote client.

    The remote HTTP server cannot write to a caller's filesystem. The attachment
    is therefore always returned as an inline FastMCP file payload.

    Args:
        message_id: The Graph message ID.
        attachment_id: The Graph attachment ID.
        profile: Ignored in HTTP mode; identity comes from the caller's token.

    Returns:
        The attachment as an inline FastMCP file payload.
    """
    return await download_attachment(
        DownloadAttachmentInput(
            message_id=params.message_id,
            attachment_id=params.attachment_id,
            profile=params.profile,
        )
    )


# ---------------------------------------------------------------------------
# save_attachment_to_path (role_obot#23)
# ---------------------------------------------------------------------------
#
# Ported from OraCool/outlook-mcp's save_attachment_to_path (tools/email_reader.py,
# backed by GraphMailClient.get_attachment_raw_bytes in auth/graph_client.py). The
# original fetches full attachment metadata (including contentBytes, discarded) via
# GET .../attachments/{id} and separately streams the binary via GET
# .../attachments/{id}/$value. This port keeps that two-call shape but trims the
# metadata call to a $select that excludes contentBytes (guinacio's own
# list_attachments does the same), and uses GraphClient.get_raw() — already present
# in graph.py for exactly this purpose — for the binary GET instead of hand-rolling a
# second httpx call.
#
# Unlike download_attachment (which can fall back to returning the file inline over
# http transport), save_attachment_to_path always requires a server-local path, so
# there is no http-mode variant: it is only registered in stdio mode (see register()
# below), matching how download_attachment itself is split by transport.


class SaveAttachmentToPathInput(ToolRequestModel):
    email_id: str
    attachment_id: str
    file_path: Path
    profile: str | None = None


async def save_attachment_to_path(
    params: SaveAttachmentToPathInput,
) -> SaveAttachmentToPathResponse:
    """
    Download an email attachment and write it directly to local disk.

    No base64 round-trip through the tool response — the binary is streamed via
    Graph's ``$value`` endpoint and written straight to ``file_path``. This is the
    reliable way to pull a PDF/DOCX/image attachment to disk; routing bytes through
    an MCP tool response hits context/token limits for anything but tiny files.

    Args:
        email_id: The Graph message ID.
        attachment_id: The Graph attachment ID.
        file_path: Exact target file path, or an existing directory (in which case
                   the attachment's own filename is used). Missing parent
                   directories are created automatically.
        profile: Microsoft 365 profile to use. Omit to use the default profile.

    Returns:
        Structured save confirmation (path, filename, size, content type), or
        success=False with an error when the attachment is not a downloadable
        file attachment (e.g. an itemAttachment/contact/event carried on the
        message, which Graph's $value endpoint does not serve as a plain file).
    """
    if get_app_config().transport == "http":
        raise ValueError(
            "save_attachment_to_path is not available in multi-user http mode "
            "(the server's disk is not the caller's disk); it is only "
            "registered in stdio mode."
        )

    g = get_graph(params.profile)
    base = f"/me/messages/{params.email_id}/attachments/{params.attachment_id}"

    raw_meta = await g.get(base, params={"$select": "id,name,size,contentType"}) or {}
    odata_type = raw_meta.get("@odata.type", "")
    if odata_type and odata_type != "#microsoft.graph.fileAttachment":
        return SaveAttachmentToPathResponse(
            success=False,
            action="save_attachment_to_path",
            email_id=params.email_id,
            attachment_id=params.attachment_id,
            error=(
                f"Unsupported attachment type {odata_type!r} — only file "
                "attachments can be saved to disk."
            ),
        )

    raw_bytes = await g.get_raw(f"{base}/$value")

    att_name = raw_meta.get("name") or "attachment"
    content_type = raw_meta.get("contentType") or "application/octet-stream"

    target = params.file_path
    if target.exists() and target.is_dir():
        safe_name = Path(att_name).name  # strip any directory components
        if not safe_name or safe_name.startswith("."):
            safe_name = "attachment"
        target = target / safe_name
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(raw_bytes)

    size_bytes = len(raw_bytes)
    return SaveAttachmentToPathResponse(
        success=True,
        action="save_attachment_to_path",
        email_id=params.email_id,
        attachment_id=params.attachment_id,
        file_path=str(target),
        filename=att_name,
        size_bytes=size_bytes,
        size_display=format_size_display(size_bytes),
        content_type=content_type,
    )


# ---------------------------------------------------------------------------
# Private helpers
# ---------------------------------------------------------------------------
def register(server) -> None:
    """Register all attachment tools with the given FastMCP server instance."""
    register_tool(server, list_attachments, annotations=READ_ONLY_TOOL)
    if get_app_config().transport == "http":
        register_tool(
            server,
            download_attachment_http,
            name="download_attachment",
            annotations=READ_ONLY_TOOL,
        )
    else:
        register_tool(server, download_attachment, annotations=WRITE_TOOL)
        register_tool(server, save_attachment_to_path, annotations=WRITE_TOOL)
