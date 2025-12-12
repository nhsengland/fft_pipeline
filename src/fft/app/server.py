"""FastHTML web interface for FFT Pipeline."""

import subprocess
import webbrowser
import logging
from pathlib import Path

from fasthtml.common import *

from fft.config import RAW_DIR, OUTPUTS_DIR, SERVICE_TYPES, FILE_PATTERNS

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Accessible CSS with dark mode support ---
CSS = Style("""
:root {
    --bg: #ffffff;
    --bg-alt: #f8f9fa;
    --text: #0b0c0c;
    --text-muted: #505a5f;
    --primary: #1d70b8;
    --primary-hover: #003078;
    --primary-text: #ffffff;
    --success: #00703c;
    --success-bg: #cce2d8;
    --error: #d4351c;
    --error-bg: #f6d7d2;
    --border: #b1b4b6;
    --focus: #ffdd00;
    --focus-text: #0b0c0c;
}

@media (prefers-color-scheme: dark) {
    :root {
        --bg: #1a1a1a;
        --bg-alt: #262626;
        --text: #f3f3f3;
        --text-muted: #b8b8b8;
        --primary: #5694ca;
        --primary-hover: #85b4dc;
        --primary-text: #0b0c0c;
        --success: #5bb98c;
        --success-bg: #1e3a2f;
        --error: #f47738;
        --error-bg: #3d2117;
        --border: #626a6e;
        --focus: #ffdd00;
        --focus-text: #0b0c0c;
    }
}

*, *::before, *::after { box-sizing: border-box; }

body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    font-size: 1.125rem;
    line-height: 1.6;
    max-width: 640px;
    margin: 0 auto;
    padding: 2rem 1rem;
    background: var(--bg);
    color: var(--text);
}

h1 {
    font-size: 2rem;
    font-weight: 700;
    margin-bottom: 1.5rem;
    border-bottom: 4px solid var(--primary);
    padding-bottom: 0.5rem;
}

label {
    display: block;
    font-weight: 600;
    margin-bottom: 0.5rem;
}

.field { margin-bottom: 1.5rem; }

select {
    width: 100%;
    padding: 0.75rem;
    font-size: 1.125rem;
    border: 2px solid var(--border);
    border-radius: 0;
    background: var(--bg);
    color: var(--text);
    appearance: none;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23505a5f' d='M1.5 4L6 8.5 10.5 4'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 0.75rem center;
}

select:focus {
    outline: 3px solid var(--focus);
    outline-offset: 0;
    box-shadow: inset 0 0 0 2px var(--focus-text);
}

.file-list {
    background: var(--bg-alt);
    border: 1px solid var(--border);
    padding: 1rem;
    margin-bottom: 1.5rem;
}

.file-list-title {
    font-weight: 600;
    margin-bottom: 0.5rem;
    color: var(--text-muted);
    font-size: 0.875rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.file-list ul {
    margin: 0;
    padding-left: 1.25rem;
}

.file-list li {
    padding: 0.25rem 0;
    font-family: monospace;
    font-size: 0.95rem;
}

.actions {
    display: flex;
    flex-wrap: wrap;
    gap: 0.75rem;
    margin-top: 1.5rem;
}

button {
    padding: 0.875rem 1.5rem;
    font-size: 1.125rem;
    font-weight: 600;
    border: none;
    border-radius: 0;
    cursor: pointer;
    transition: background 0.15s;
}

button:focus {
    outline: 3px solid var(--focus);
    outline-offset: 0;
    box-shadow: inset 0 0 0 2px var(--focus-text);
}

.btn-primary {
    background: var(--primary);
    color: var(--primary-text);
}

.btn-primary:hover { background: var(--primary-hover); }

.btn-secondary {
    background: var(--bg-alt);
    color: var(--text);
    border: 2px solid var(--border);
}

.btn-secondary:hover { background: var(--border); }

.status {
    margin-top: 2rem;
    padding: 1rem;
    border-left: 4px solid var(--border);
    background: var(--bg-alt);
}

.status.success {
    border-color: var(--success);
    background: var(--success-bg);
}

.status.error {
    border-color: var(--error);
    background: var(--error-bg);
}

.status-title {
    font-weight: 700;
    font-size: 1.125rem;
    margin-bottom: 0.5rem;
}

.log-output {
    font-family: monospace;
    font-size: 0.875rem;
    line-height: 1.4;
    white-space: pre-wrap;
    max-height: 250px;
    overflow-y: auto;
    background: var(--text);
    color: var(--bg);
    padding: 1rem;
    margin-top: 0.75rem;
}

details summary {
    cursor: pointer;
    font-weight: 600;
    color: var(--primary);
}

details summary:focus {
    outline: 3px solid var(--focus);
    outline-offset: 2px;
}

.visually-hidden {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    border: 0;
}
""")

app, rt = fast_app(hdrs=[CSS])


# --- Helpers ---
def get_raw_files(service_type: str = None) -> list[Path]:
    if not RAW_DIR.exists():
        return []
    pattern = FILE_PATTERNS.get(service_type, "*.xlsx") if service_type else "*.xlsx"
    return sorted(RAW_DIR.glob(pattern), reverse=True)


def get_months(service_type: str) -> list[str]:
    return sorted({f.stem.split()[-1] for f in get_raw_files(service_type)}, reverse=True)


def run_cmd(service: str, month: str) -> tuple[bool, str]:
    flag = [k for k, v in SERVICE_TYPES.items() if v == service][0]
    cmd = ["uv", "run", "python", "-m", "src.fft", f"--{flag}"]
    if month and month != "all":
        cmd.extend(["--month", month])

    project_root = Path(__file__).parent.parent
    logger.info(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=project_root)
    logger.info(f"Return code: {result.returncode}")

    return result.returncode == 0, result.stdout if result.returncode == 0 else (
        result.stderr or result.stdout
    )


# --- Components ---
def service_select():
    opts = [Option("-- Select service type --", value="", disabled=True, selected=True)]
    opts += [Option(name.title(), value=name) for _, name in SERVICE_TYPES.items()]
    return Select(
        *opts,
        name="service",
        id="service",
        hx_get="/months",
        hx_target="#month-container",
        hx_trigger="change",
        aria_required="true",
    )


def month_select(months=None):
    opts = [Option("All months", value="all")]
    if months:
        opts += [Option(m, value=m) for m in months]
    return Select(*opts, name="month", id="month")


def file_list_box(files):
    if not files:
        return Div(
            Div("Available files", cls="file-list-title"),
            P("No files found in raw folder"),
            cls="file-list",
        )
    items = [Li(f.name) for f in files[:8]]
    extra = Li(f"... and {len(files) - 8} more") if len(files) > 8 else None
    return Div(
        Div("Available files", cls="file-list-title"),
        Ul(*items, extra) if extra else Ul(*items),
        cls="file-list",
    )


def status_box(success: bool, msg: str, log: str = None):
    cls = "status success" if success else "status error"
    icon = "✓" if success else "✗"
    content = [Div(f"{icon} {msg}", cls="status-title")]
    if log:
        content.append(
            Details(
                Summary("View log output"),
                Div(log, cls="log-output", role="log", aria_live="polite"),
            )
        )
    return Div(*content, cls=cls, role="alert", aria_live="polite")


# --- Routes ---
@rt("/")
def get():
    return Titled(
        "FFT Pipeline",
        Form(
            Div(Label("Service Type", for_="service"), service_select(), cls="field"),
            Div(
                Label("Month", for_="month"),
                Div(month_select(), id="month-container"),
                cls="field",
            ),
            Div(id="file-list", hx_get="/files", hx_trigger="load"),
            Div(
                Button("Run Pipeline", type="submit", cls="btn-primary"),
                Button(
                    "Open Output Folder",
                    type="button",
                    cls="btn-secondary",
                    hx_post="/open-output",
                    hx_swap="none",
                ),
                cls="actions",
            ),
            Div(id="status", aria_live="polite"),
            hx_post="/run",
            hx_target="#status",
        ),
    )


@rt("/months")
def get(service: str = ""):
    return month_select(get_months(service) if service else None)


@rt("/files")
def get(service: str = ""):
    return file_list_box(get_raw_files(service or None))


@rt("/run")
async def post(service: str, month: str):
    if not service:
        return status_box(False, "Please select a service type")
    success, log = run_cmd(service, month)
    return status_box(
        success, "Pipeline completed successfully" if success else "Pipeline failed", log
    )


@rt("/open-output")
def post():
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    webbrowser.open(f"file://{OUTPUTS_DIR.absolute()}")
    return ""


if __name__ == "__main__":
    serve()
