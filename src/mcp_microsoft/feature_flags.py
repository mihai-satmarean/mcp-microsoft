"""Feature flags for optional mcp-microsoft services.

Environment flags are authoritative when present. When a flag is absent, the
server falls back to account-type detection so manual installs with corporate
profiles can auto-enable Teams and SharePoint.

"""

from __future__ import annotations

from mcp_microsoft.config import AppConfig, env_flag, get_app_config

_EXPLICIT_ONLY_FLAGS = frozenset(
    {
        "MCP_ENABLE_TEAMS_MEETING_ARTIFACTS",
        "MCP_ENABLE_TEAMS_AI_INSIGHTS",
    }
)


def _resolve_profile_for_detection(
    profile_name: str | None,
    config: AppConfig | None = None,
):
    from mcp_microsoft.profiles import get_profile_manager

    return get_profile_manager(config=config).resolve_profile(profile_name)


def _get_configured_flag(
    env_name: str,
    config: AppConfig | None = None,
) -> bool | None:
    runtime_config = config or get_app_config()
    if env_name == "MCP_ENABLE_TEAMS":
        return runtime_config.enable_teams
    if env_name == "MCP_ENABLE_SHAREPOINT":
        return runtime_config.enable_sharepoint
    if env_name == "MCP_ENABLE_TEAMS_MEETING_ARTIFACTS":
        return runtime_config.enable_teams_meeting_artifacts
    if env_name == "MCP_ENABLE_TEAMS_AI_INSIGHTS":
        return runtime_config.enable_teams_ai_insights
    return env_flag(env_name)


def resolve_optional_service_enabled(
    env_name: str,
    profile_name: str | None = None,
    config: AppConfig | None = None,
) -> bool:
    """Resolve an optional service flag with account-type fallback.

    If the env var is set, respect it even when the value is falsy.
    If it is absent, fall back to corporate-account detection for the resolved
    profile. If no profile is available yet, fail safe to False.
    """
    explicit = _get_configured_flag(env_name, config=config)
    if env_name in _EXPLICIT_ONLY_FLAGS:
        return bool(explicit)
    if explicit is not None:
        return explicit

    # http (multi-user remote) mode has no single default profile, so the
    # corporate-account heuristic can't apply — flags must be explicit there.
    runtime_config = config or get_app_config()
    if runtime_config.transport == "http":
        return False

    try:
        from mcp_microsoft.profiles import is_corporate_account

        profile = _resolve_profile_for_detection(profile_name, config=config)
        return is_corporate_account(profile)
    except Exception:
        return False


def is_teams_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Teams should be enabled for the active/default profile."""
    return resolve_optional_service_enabled("MCP_ENABLE_TEAMS", config=config)


def is_sharepoint_enabled(config: AppConfig | None = None) -> bool:
    """Return True when SharePoint should be enabled for the active/default profile."""
    return resolve_optional_service_enabled("MCP_ENABLE_SHAREPOINT", config=config)


def is_teams_meeting_artifacts_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Teams meeting transcript/recording tools should be enabled."""
    runtime_config = config or get_app_config()
    if not is_teams_enabled(config=runtime_config):
        return False
    return bool(runtime_config.enable_teams_meeting_artifacts)


def is_teams_ai_insights_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Teams Copilot AI insight tools should be enabled."""
    runtime_config = config or get_app_config()
    if not is_teams_enabled(config=runtime_config):
        return False
    return bool(runtime_config.enable_teams_ai_insights)


def is_file_upload_enabled(config: AppConfig | None = None) -> bool:
    """Return True when the context-free file-upload app should be registered.

    Tri-state resolution mirroring the other optional-service flags, but with a
    transport-aware default rather than a corporate-account fallback: an explicit
    ``MCP_ENABLE_FILE_UPLOAD`` (truthy or falsy) always wins; when unset the
    feature is ON in http mode (remote callers have no local disk to pass via
    ``local_path``) and OFF in stdio mode (local users already have ``local_path``).
    """
    runtime_config = config or get_app_config()
    explicit = runtime_config.enable_file_upload
    if explicit is not None:
        return explicit
    return runtime_config.transport == "http"


def resolve_upload_max_bytes(config: AppConfig | None = None) -> int:
    """Return the per-file upload size limit in bytes, validating it is positive.

    Raises:
        ValueError: if ``upload_max_mb`` is not a positive integer.
    """
    runtime_config = config or get_app_config()
    mb = runtime_config.upload_max_mb
    if mb <= 0:
        raise ValueError(
            f"MCP_UPLOAD_MAX_MB must be a positive integer (got {mb})."
        )
    return mb * 1024 * 1024


def resolve_upload_global_budget_bytes(config: AppConfig | None = None) -> int:
    """Return the global encoded-byte upload budget in bytes, validating positivity.

    Bounds the total base64 footprint of all uploaded files across every scope
    held in process memory (``uploads.ScopedFileUpload`` enforces it). Consistent
    with :func:`resolve_upload_max_bytes`: rejects a non-positive value.

    Raises:
        ValueError: if ``upload_global_budget_mb`` is not a positive integer.
    """
    runtime_config = config or get_app_config()
    mb = runtime_config.upload_global_budget_mb
    if mb <= 0:
        raise ValueError(
            f"MCP_UPLOAD_GLOBAL_BUDGET_MB must be a positive integer (got {mb})."
        )
    return mb * 1024 * 1024


def is_todo_enabled(config: AppConfig | None = None) -> bool:
    """Return True when Microsoft To Do (Tasks) tools should be registered.

    role_obot#23. Unlike Teams/SharePoint this defaults ON in both transports
    (no corporate-account detection) — To Do is a standard mailbox feature,
    not an optional add-on service. It does add one new delegated scope
    (``Tasks.ReadWrite``, see ``profiles.TODO_SCOPES``) that is requested at
    consent time whenever this flag resolves True. Set ``MCP_ENABLE_TODO=false``
    to omit the tool group (and the scope request) entirely.
    """
    runtime_config = config or get_app_config()
    explicit = runtime_config.enable_todo
    if explicit is not None:
        return explicit
    return True


def is_delta_enabled(config: AppConfig | None = None) -> bool:
    """Return True when delta-sync, batch-triage, and digest tools should be registered.

    role_obot#23. Defaults ON for the same reason as :func:`is_todo_enabled` —
    delta queries read the same mailbox/calendar/contacts a caller can already
    list. Set ``MCP_ENABLE_DELTA=false`` to omit the tool group entirely.
    """
    runtime_config = config or get_app_config()
    explicit = runtime_config.enable_delta
    if explicit is not None:
        return explicit
    return True


def is_deletion_disabled(config: AppConfig | None = None) -> bool:
    """Return True when destructive hard-delete tools should NOT be registered.

    Controlled by ``MCP_DISABLE_DELETION_TOOLS``. When truthy, the server
    omits all permanent-deletion tools (e.g. delete_email, delete_event,
    delete_drive_item, remove_ms_profile). Recoverable variants such as
    trash_email remain available.
    """
    runtime_config = config or get_app_config()
    return bool(runtime_config.disable_deletion_tools)
