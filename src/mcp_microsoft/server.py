"""
mcp-microsoft FastMCP server.

Registers all Microsoft 365 tools (Mail, Calendar, OneDrive, etc.) from the
tools/ submodules. Entry point: mcp-microsoft (see pyproject.toml).

Run:
    python -m mcp_microsoft.server
    # or via the installed script:
    mcp-microsoft
"""

from __future__ import annotations

import logging
import os
import sys
from contextlib import asynccontextmanager
from typing import Any

from fastmcp import FastMCP

from mcp_microsoft.config import AppConfig, get_app_config, validate_http_config
from mcp_microsoft.feature_flags import (
    is_delta_enabled,
    is_file_upload_enabled,
    is_sharepoint_enabled,
    is_teams_ai_insights_enabled,
    is_teams_enabled,
    is_teams_meeting_artifacts_enabled,
    is_todo_enabled,
    resolve_upload_global_budget_bytes,
    resolve_upload_max_bytes,
)
from mcp_microsoft.graph import close_http_clients, initialize_http_clients
from mcp_microsoft.profiles import get_profile_manager

_log = logging.getLogger(__name__)
_mcp_server: FastMCP | None = None

# Microsoft Graph resource prefix for delegated permission URIs advertised to
# Azure via AzureProvider.additional_authorize_scopes.
_GRAPH_RESOURCE = "https://graph.microsoft.com"


def build_graph_authorize_scopes(config: AppConfig) -> list[str]:
    """Build the delegated Graph scopes for ``additional_authorize_scopes``.

    Mirrors ``profiles.build_default_scopes`` semantics but emits the full
    ``https://graph.microsoft.com/<Scope>`` URI form that AzureProvider forwards
    upstream. ``DEFAULT_SCOPES`` are always present; Teams/SharePoint/To Do
    scopes are added only when their feature flag resolves True (in http mode
    the flags are env-only — no corporate-profile fallback). ``offline_access``
    is appended unprefixed because it is an OIDC scope, not a Graph resource
    scope.
    """
    from mcp_microsoft.profiles import (
        DEFAULT_SCOPES,
        SHAREPOINT_SCOPES,
        TEAMS_AI_INSIGHT_SCOPES,
        TEAMS_MEETING_ARTIFACT_SCOPES,
        TEAMS_SCOPES,
        TODO_SCOPES,
    )

    graph_scopes: list[str] = list(DEFAULT_SCOPES)
    if is_teams_enabled(config=config):
        graph_scopes.extend(TEAMS_SCOPES)
        if is_teams_meeting_artifacts_enabled(config=config):
            graph_scopes.extend(TEAMS_MEETING_ARTIFACT_SCOPES)
        if is_teams_ai_insights_enabled(config=config):
            graph_scopes.extend(TEAMS_AI_INSIGHT_SCOPES)
    if is_sharepoint_enabled(config=config):
        graph_scopes.extend(SHAREPOINT_SCOPES)
    if is_todo_enabled(config=config):
        graph_scopes.extend(TODO_SCOPES)

    deduped = list(dict.fromkeys(graph_scopes))
    scopes = [f"{_GRAPH_RESOURCE}/{scope}" for scope in deduped]
    scopes.append("offline_access")
    return scopes


def _build_http_middleware(config: AppConfig) -> list[Any]:
    """Build the middleware stack for http (multi-user) transport.

    Order matters: rate limiting runs first so throttled requests are
    rejected before audit logging (and tool execution) ever sees them. Audit
    logging always runs so every authenticated tool call that gets through
    is recorded. Metrics recording runs last (innermost) so it wraps the tool
    call most tightly and times the tool itself; it is always present so the
    observability routes have data even before a token is configured.
    """
    from mcp_microsoft.middleware import (
        AuditLoggingMiddleware,
        MetricsMiddleware,
        UserRateLimitMiddleware,
    )

    stack: list[Any] = []
    if config.rate_limit_rps > 0:
        stack.append(
            UserRateLimitMiddleware(max_requests_per_second=config.rate_limit_rps)
        )
    stack.append(AuditLoggingMiddleware())
    stack.append(MetricsMiddleware())
    return stack


def _register_health_route(mcp: FastMCP) -> None:
    """Register an unauthenticated ``GET /health`` for load balancers.

    ``@custom_route`` handlers are mounted outside FastMCP's
    ``RequireAuthMiddleware`` (which wraps only the MCP endpoint route
    itself), so this stays reachable without a bearer token — exactly what a
    health check needs.
    """
    from starlette.requests import Request
    from starlette.responses import JSONResponse, Response

    @mcp.custom_route("/health", methods=["GET"])
    async def health(_request: Request) -> Response:
        return JSONResponse({"status": "ok", "transport": "http"})


# WWW-Authenticate challenge advertising both accepted credential schemes for
# the observability routes. Basic is listed first so a browser hitting
# /dashboard shows a native login box; Bearer lets scrapers/curl use a token.
_STATS_CHALLENGE = 'Basic realm="mcp-microsoft stats", Bearer'


def _stats_authorized(request: Any, token: str) -> bool:
    """Return True if *request* carries the configured stats token.

    Accepts either ``Authorization: Bearer <token>`` or HTTP Basic where the
    password equals the token (any username). The comparison is timing-safe
    (``hmac.compare_digest`` over utf-8 bytes) and the token is never logged.
    Any malformed header yields False rather than raising.
    """
    import base64
    import binascii
    import hmac

    expected = token.encode("utf-8")
    header = request.headers.get("Authorization", "")
    scheme, _, credentials = header.partition(" ")
    scheme = scheme.strip().lower()
    credentials = credentials.strip()

    if scheme == "bearer":
        return hmac.compare_digest(credentials.encode("utf-8"), expected)
    if scheme == "basic":
        try:
            decoded = base64.b64decode(credentials, validate=True).decode("utf-8")
        except (binascii.Error, ValueError, UnicodeDecodeError):
            return False
        _user, sep, password = decoded.partition(":")
        if not sep:
            return False
        return hmac.compare_digest(password.encode("utf-8"), expected)
    return False


def _register_stats_routes(mcp: FastMCP, config: AppConfig) -> None:
    """Register the token-gated observability routes (http mode only).

    Three routes served by the same server: ``GET /metrics`` (Prometheus text),
    ``GET /stats`` (JSON snapshot), and ``GET /dashboard`` (a self-contained
    HTML page). All require the ``MCP_STATS_TOKEN`` credential; the caller only
    reaches here when that token is non-empty. ``/health`` is unaffected.
    """
    from starlette.requests import Request
    from starlette.responses import (
        HTMLResponse,
        JSONResponse,
        PlainTextResponse,
        Response,
    )

    from mcp_microsoft.metrics import get_metrics_registry

    # Captured once at registration; the token is never placed in a log line.
    token = config.stats_token

    # These responses carry live operational data (and, on the dashboard, the
    # page a browser holds behind Basic auth). Forbid any shared or browser
    # cache from retaining them so a scrape/response is never served from cache
    # to a later, differently-authorized reader.
    _NO_STORE = {"Cache-Control": "no-store"}

    def _challenge() -> Response:
        return Response(
            status_code=401,
            headers={"WWW-Authenticate": _STATS_CHALLENGE},
        )

    @mcp.custom_route("/metrics", methods=["GET"], include_in_schema=False)
    async def metrics_route(request: Request) -> Response:
        if not _stats_authorized(request, token):
            return _challenge()
        body = get_metrics_registry().render_prometheus()
        return PlainTextResponse(
            body,
            media_type="text/plain; version=0.0.4; charset=utf-8",
            headers=_NO_STORE,
        )

    @mcp.custom_route("/stats", methods=["GET"], include_in_schema=False)
    async def stats_route(request: Request) -> Response:
        if not _stats_authorized(request, token):
            return _challenge()
        return JSONResponse(get_metrics_registry().snapshot(), headers=_NO_STORE)

    @mcp.custom_route("/dashboard", methods=["GET"], include_in_schema=False)
    async def dashboard_route(request: Request) -> Response:
        if not _stats_authorized(request, token):
            return _challenge()
        return HTMLResponse(_DASHBOARD_HTML, headers=_NO_STORE)


# Self-contained observability dashboard: inline CSS + JS, no external requests
# (a strict environment / CSP-conscious operator can serve it as-is). It polls
# the same-origin /stats endpoint every 10s; the browser re-sends the Basic
# credentials it prompted for on the initial /dashboard load automatically, so
# no token handling lives in this page.
_DASHBOARD_HTML = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>mcp-microsoft — observability</title>
<style>
  :root {
    --bg: #f6f7f9; --card: #ffffff; --fg: #1b1f24; --muted: #6b7280;
    --border: #e5e7eb; --accent: #2563eb; --err: #dc2626; --ok: #16a34a;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #0d1117; --card: #161b22; --fg: #e6edf3; --muted: #8b949e;
      --border: #30363d; --accent: #58a6ff; --err: #f85149; --ok: #3fb950;
    }
  }
  * { box-sizing: border-box; }
  body { margin: 0; padding: 24px; background: var(--bg); color: var(--fg);
    font: 14px/1.5 system-ui, -apple-system, Segoe UI, Roboto, sans-serif; }
  h1 { font-size: 18px; margin: 0 0 4px; }
  .sub { color: var(--muted); font-size: 12px; margin-bottom: 20px; }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
    gap: 12px; margin-bottom: 20px; }
  .card { background: var(--card); border: 1px solid var(--border);
    border-radius: 10px; padding: 14px 16px; }
  .card .label { color: var(--muted); font-size: 11px; text-transform: uppercase;
    letter-spacing: .04em; }
  .card .value { font-size: 24px; font-weight: 600; margin-top: 4px; }
  section { background: var(--card); border: 1px solid var(--border);
    border-radius: 10px; padding: 16px; margin-bottom: 20px; }
  section h2 { font-size: 14px; margin: 0 0 12px; }
  svg { width: 100%; height: 80px; display: block; }
  .tbl-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 13px;
    min-width: 520px; }
  th, td { text-align: right; padding: 6px 10px; border-bottom: 1px solid var(--border);
    white-space: nowrap; }
  th:first-child, td:first-child { text-align: left; }
  th { color: var(--muted); font-weight: 600; cursor: pointer; user-select: none; }
  td.mono, .oid { font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
    font-size: 12px; color: var(--muted); }
  .err { color: var(--err); }
  .empty { color: var(--muted); text-align: center; padding: 16px; }
  #status { color: var(--muted); font-size: 12px; }
</style>
</head>
<body>
  <h1>mcp-microsoft &mdash; observability</h1>
  <div class="sub">http transport &middot; single worker &middot; metrics reset on restart &middot;
    <span id="status">loading&hellip;</span></div>

  <div class="cards" id="cards"></div>

  <section>
    <h2>Traffic &mdash; calls per minute (last 60m)</h2>
    <svg id="spark" viewBox="0 0 600 80" preserveAspectRatio="none"></svg>
  </section>

  <section>
    <h2>Tools</h2>
    <div class="tbl-wrap">
      <table id="tools">
        <thead><tr>
          <th data-k="name">Tool</th><th data-k="calls">Calls</th>
          <th data-k="errors">Errors</th><th data-k="avg_ms">Avg ms</th>
          <th data-k="p50_ms">p50 ms</th><th data-k="p95_ms">p95 ms</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </section>

  <section>
    <h2>Users <span id="usermeta" class="sub"></span></h2>
    <div class="tbl-wrap">
      <table id="users">
        <thead><tr>
          <th data-k="username">User</th><th data-k="oid">OID</th>
          <th data-k="calls">Calls</th><th data-k="errors">Errors</th>
          <th data-k="last_seen_iso">Last seen</th>
        </tr></thead>
        <tbody></tbody>
      </table>
    </div>
  </section>

<script>
  var toolSort = { k: "calls", dir: -1 };
  var userSort = { k: "last_seen_iso", dir: -1 };

  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }
  function fmtDur(s) {
    s = Math.floor(s);
    var d = Math.floor(s / 86400); s %= 86400;
    var h = Math.floor(s / 3600); s %= 3600;
    var m = Math.floor(s / 60); s %= 60;
    if (d) return d + "d " + h + "h";
    if (h) return h + "h " + m + "m";
    if (m) return m + "m " + s + "s";
    return s + "s";
  }
  function pct(n, d) { return d ? (100 * n / d).toFixed(1) + "%" : "0.0%"; }

  function card(label, value) {
    return '<div class="card"><div class="label">' + esc(label) +
      '</div><div class="value">' + esc(value) + "</div></div>";
  }

  function sparkline(series) {
    var svg = document.getElementById("spark");
    var w = 600, h = 80, n = series.length;
    var max = 1;
    series.forEach(function (p) { if (p.calls > max) max = p.calls; });
    var pts = series.map(function (p, i) {
      var x = n > 1 ? (i / (n - 1)) * w : 0;
      var y = h - (p.calls / max) * (h - 6) - 3;
      return x.toFixed(1) + "," + y.toFixed(1);
    }).join(" ");
    var stroke = getComputedStyle(document.documentElement)
      .getPropertyValue("--accent").trim() || "#2563eb";
    svg.innerHTML =
      '<polyline fill="none" stroke="' + stroke + '" stroke-width="2" points="' +
      pts + '"/>';
  }

  function sortRows(rows, s) {
    return rows.slice().sort(function (a, b) {
      var x = a[s.k], y = b[s.k];
      if (typeof x === "number" && typeof y === "number") return (x - y) * s.dir;
      return String(x).localeCompare(String(y)) * s.dir;
    });
  }

  function renderTools(tools) {
    var body = document.querySelector("#tools tbody");
    if (!tools.length) { body.innerHTML = '<tr><td colspan="6" class="empty">No calls yet.</td></tr>'; return; }
    body.innerHTML = sortRows(tools, toolSort).map(function (t) {
      return "<tr><td>" + esc(t.name) + "</td><td>" + t.calls +
        '</td><td class="' + (t.errors ? "err" : "") + '">' + t.errors +
        "</td><td>" + t.avg_ms + "</td><td>" + t.p50_ms + "</td><td>" +
        t.p95_ms + "</td></tr>";
    }).join("");
  }

  function renderUsers(u) {
    document.getElementById("usermeta").textContent =
      "tracked " + u.count + (u.evicted ? " · evicted " + u.evicted : "");
    var body = document.querySelector("#users tbody");
    if (!u.top.length) { body.innerHTML = '<tr><td colspan="5" class="empty">No users yet.</td></tr>'; return; }
    body.innerHTML = sortRows(u.top, userSort).map(function (r) {
      return "<tr><td>" + esc(r.username) + '</td><td class="oid">' + esc(r.oid) +
        "</td><td>" + r.calls + '</td><td class="' + (r.errors ? "err" : "") +
        '">' + r.errors + '</td><td class="mono">' +
        esc((r.last_seen_iso || "").replace("T", " ").slice(0, 19)) + "</td></tr>";
    }).join("");
  }

  function render(d) {
    var s = d.server, tr = d.traffic;
    document.getElementById("cards").innerHTML =
      card("Uptime", fmtDur(s.uptime_s)) +
      card("Total calls", s.total_calls) +
      card("Error rate", pct(s.total_errors, s.total_calls)) +
      card("Users tracked", d.users.count) +
      card("Unknown tools", s.unknown_tool_calls) +
      card("Calls (5m)", tr.last_5m.calls) +
      card("Calls (60m)", tr.last_60m.calls);
    sparkline(tr.per_minute);
    renderTools(d.tools);
    renderUsers(d.users);
    document.getElementById("status").textContent =
      "updated " + new Date().toLocaleTimeString();
  }

  function bindSort(tableId, state, rerender) {
    document.querySelectorAll("#" + tableId + " th").forEach(function (th) {
      th.addEventListener("click", function () {
        var k = th.getAttribute("data-k");
        if (state.k === k) state.dir *= -1; else { state.k = k; state.dir = -1; }
        rerender();
      });
    });
  }

  var last = null;
  bindSort("tools", toolSort, function () { if (last) renderTools(last.tools); });
  bindSort("users", userSort, function () { if (last) renderUsers(last.users); });

  function poll() {
    fetch("/stats", { headers: { "Accept": "application/json" } })
      .then(function (r) {
        if (!r.ok) throw new Error("HTTP " + r.status);
        return r.json();
      })
      .then(function (d) { last = d; render(d); })
      .catch(function (e) {
        document.getElementById("status").textContent = "error: " + e.message;
      });
  }
  poll();
  setInterval(poll, 10000);
</script>
</body>
</html>
"""


def _build_azure_provider(config: AppConfig):
    """Construct the FastMCP AzureProvider for http (multi-user) transport.

    Imported lazily so stdio mode never touches ``fastmcp[azure]``. The
    constructor performs no network I/O (JWKS/OBO are fetched lazily at request
    time), so it is safe to build during server construction and in tests.
    """
    from fastmcp.server.auth.providers.azure import AzureProvider

    return AzureProvider(
        client_id=config.auth_client_id,
        client_secret=config.auth_client_secret,
        tenant_id=config.auth_tenant_id,
        base_url=config.base_url,
        required_scopes=[config.auth_required_scope],
        additional_authorize_scopes=build_graph_authorize_scopes(config),
    )


@asynccontextmanager
async def app_lifespan(_server: FastMCP):
    """Initialize shared runtime resources once for the server process."""
    await initialize_http_clients()
    try:
        yield {}
    finally:
        await close_http_clients()


def create_mcp_server(config: AppConfig | None = None) -> FastMCP:
    """Build a FastMCP server using the current runtime configuration."""
    if config is not None:
        # Programmatic-embedding path: make the global cache the SAME object we
        # thread below, so the ~11 sites that read get_app_config() directly
        # (disk-tool gates, runtime path rejections, graph.get_graph dispatch)
        # can't disagree with this explicit config. See config.set_app_config.
        from mcp_microsoft.config import set_app_config

        set_app_config(config)
    runtime_config = config or get_app_config()
    http_mode = runtime_config.transport == "http"

    if not http_mode:
        # http mode never touches ProfileManager — identity always comes from
        # the caller's bearer token via OboTokenProvider (see graph.get_graph
        # and feature_flags.resolve_optional_service_enabled). Warming it up
        # here would only create ~/.microsoft-mcp on a server that will never
        # read or write it.
        get_profile_manager(config=runtime_config)

    auth = _build_azure_provider(runtime_config) if http_mode else None
    middleware = _build_http_middleware(runtime_config) if http_mode else None
    mcp = FastMCP(
        "mcp-microsoft",
        lifespan=app_lifespan,
        auth=auth,
        middleware=middleware,
        mask_error_details=True if http_mode else None,
    )

    if http_mode:
        _register_health_route(mcp)
        if runtime_config.stats_token:
            _register_stats_routes(mcp, runtime_config)
        else:
            _log.info(
                "Observability routes not registered (MCP_STATS_TOKEN unset; "
                "/metrics, /stats, /dashboard disabled)"
            )

    from mcp_microsoft.tools import attachments
    from mcp_microsoft.tools import batch
    from mcp_microsoft.tools import calendar
    from mcp_microsoft.tools import contacts
    from mcp_microsoft.tools import delta
    from mcp_microsoft.tools import digest
    from mcp_microsoft.tools import drafts
    from mcp_microsoft.tools import folders
    from mcp_microsoft.tools import mail
    from mcp_microsoft.tools import onedrive
    from mcp_microsoft.tools import profiles
    from mcp_microsoft.tools import services
    from mcp_microsoft.tools import sharepoint
    from mcp_microsoft.tools import todo

    mail.register(mcp)
    drafts.register(mcp)
    folders.register(mcp)
    attachments.register(mcp)
    calendar.register(mcp)
    onedrive.register(mcp)
    if http_mode:
        _log.info(
            "Profile-management tools not registered "
            "(http transport; identity is per-request from the bearer token)"
        )
    else:
        profiles.register(mcp)
    contacts.register(mcp)

    if is_sharepoint_enabled(config=runtime_config):
        sharepoint.register(mcp)
    else:
        _log.info("SharePoint tools not registered (service disabled)")

    if is_teams_enabled(config=runtime_config):
        from mcp_microsoft.tools import teams

        teams.register(mcp)
    else:
        _log.info("Teams tools not registered (service disabled)")

    # role_obot#23: To Do (Tasks) and delta/batch/digest tool groups.
    if is_todo_enabled(config=runtime_config):
        todo.register(mcp)
    else:
        _log.info("To Do tools not registered (service disabled)")

    if is_delta_enabled(config=runtime_config):
        delta.register(mcp)
        batch.register(mcp)
        digest.register(mcp)
    else:
        _log.info("Delta/batch/digest tools not registered (service disabled)")

    services.register(mcp)

    _register_file_upload(mcp, runtime_config)

    return mcp


def _register_file_upload(mcp: FastMCP, config: AppConfig) -> None:
    """Wire the context-free file-upload app when the feature resolves enabled.

    Constructs a bounded, per-caller-scoped :class:`ScopedFileUpload`, installs
    it as the process-wide singleton the Graph upload tools resolve against, and
    adds it as a provider under the default (empty) namespace — its tool names
    (``store_files``/``list_files``/``read_file``/``file_manager``) do not
    collide with any of our ~95 registered tools. Its tools flow through the same
    http middleware stack (rate limiting, audit logging, metrics) as every other
    tool, since middleware wraps ``tools/call`` dispatch, not registration.

    Always logs one info line (enabled or not). Clears the singleton when the
    feature is off so a rebuilt server never leaves a stale provider installed.
    """
    from mcp_microsoft.uploads import set_upload_provider

    if not is_file_upload_enabled(config=config):
        set_upload_provider(None)
        _log.info(
            "File-upload app not enabled (MCP_ENABLE_FILE_UPLOAD unset/false; "
            "default is on in http mode, off in stdio)"
        )
        return

    from mcp_microsoft.uploads import ScopedFileUpload

    max_bytes = resolve_upload_max_bytes(config)
    global_budget = resolve_upload_global_budget_bytes(config)
    provider = ScopedFileUpload(
        max_file_size=max_bytes, global_budget_bytes=global_budget
    )
    set_upload_provider(provider)
    mcp.add_provider(provider)
    _log.info(
        "File-upload app enabled (drag-drop UI; per-file limit %d MB, "
        "per-user quota bounded, global budget %d MB)",
        config.upload_max_mb,
        config.upload_global_budget_mb,
    )


def get_mcp_server(reset: bool = False, config: AppConfig | None = None) -> FastMCP:
    """Return the lazily constructed FastMCP server."""
    global _mcp_server
    if reset or _mcp_server is None or config is not None:
        _mcp_server = create_mcp_server(config=config)
    return _mcp_server


def reset_mcp_server() -> None:
    """Drop the cached FastMCP server so it is rebuilt on next access."""
    global _mcp_server
    _mcp_server = None


class _MCPProxy:
    """Compatibility proxy for code that imports ``server.mcp``."""

    def __getattr__(self, name: str) -> Any:
        return getattr(get_mcp_server(), name)


mcp = _MCPProxy()


def _configure_logging(config: AppConfig) -> None:
    """Configure process-wide logging for the CLI entry point.

    Without this, Python's root logger stays at WARNING with only the
    last-resort handler — which silently swallows the http-mode audit trail
    (``AuditLoggingMiddleware`` logs at INFO) and every informational
    registration line. Defaults: INFO in http mode (the audit log is a core
    feature there), WARNING in stdio (matches the previous behavior of a
    local, single-user subprocess). ``MCP_LOG_LEVEL`` overrides either.

    Logs go to stderr explicitly: in stdio transport stdout carries the MCP
    protocol stream and must never receive log lines. ``force=True`` so this
    wins even if an imported library already installed a root handler.

    Called only from :func:`main` — programmatic embedders keep full control
    of their own logging configuration.
    """
    level_name = os.environ.get("MCP_LOG_LEVEL", "").strip().upper()
    level: int | None = None
    if level_name:
        candidate = logging.getLevelName(level_name)
        if isinstance(candidate, int):
            level = candidate
    if level is None:
        level = logging.INFO if config.transport == "http" else logging.WARNING

    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
        stream=sys.stderr,
        force=True,
    )


def main() -> None:
    """CLI entry point — starts the MCP server over stdio or Streamable HTTP."""
    config = get_app_config()
    _configure_logging(config)
    if config.transport not in ("stdio", "http"):
        raise SystemExit(
            "Cannot start mcp-microsoft — MCP_TRANSPORT must be 'stdio' or "
            f"'http' (got {config.transport!r})"
        )
    if config.transport == "http":
        problems = validate_http_config(config)
        if problems:
            raise SystemExit(
                "Cannot start http transport — fix the following configuration "
                "problems:\n  - " + "\n  - ".join(problems)
            )
        get_mcp_server().run(
            transport="http",
            host=config.http_host,
            port=config.http_port,
            stateless_http=config.http_stateless,
        )
    else:
        get_mcp_server().run()


if __name__ == "__main__":
    main()
