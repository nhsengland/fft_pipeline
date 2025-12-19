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
    font-size: clamp(1.5rem, 3vw, 1.875rem);  /* 24px-30px responsive - LARGER for VI */
    line-height: 1.6;
    max-width: min(90vw, 1200px);
    margin: 0 auto;
    padding: clamp(1.5rem, 4vw, 3rem) clamp(1rem, 3vw, 2rem);  /* Responsive padding */
    background: var(--bg);
    color: var(--text);
}

h1 {
    font-size: clamp(2rem, 5vw, 2.75rem);  /* 32px-44px responsive */
    font-weight: 700;
    margin-bottom: clamp(1.5rem, 3vw, 2rem);
    border-bottom: 4px solid var(--primary);
    padding-bottom: 0.75rem;
    color: var(--text);
    background: var(--bg);
    letter-spacing: -0.02em;
}

label {
    display: block;
    font-weight: 600;
    font-size: clamp(1.375rem, 2.5vw, 1.625rem);  /* 22px-26px - larger for VI */
    margin-bottom: 0.75rem;
    color: var(--text);
    letter-spacing: 0.01em;  /* Slight letter spacing for readability */
}

.field {
    margin-bottom: clamp(1.5rem, 3vw, 2rem);
    position: relative;
}

/* Responsive form grid layout */
.form-grid {
    display: grid;
    gap: clamp(1rem, 3vw, 2rem);
    margin-bottom: clamp(1.5rem, 3vw, 2.5rem);
}

/* Mobile: single column */
@media (max-width: 767px) {
    .form-grid {
        grid-template-columns: 1fr;
    }

    /* Add subtle visual separation on mobile */
    .field:not(:last-child):after {
        content: "";
        position: absolute;
        bottom: -0.75rem;
        left: 0;
        right: 0;
        height: 1px;
        background: linear-gradient(90deg, var(--border), transparent);
        opacity: 0.3;
    }
}

/* Desktop: side-by-side fields */
@media (min-width: 768px) {
    .form-grid {
        grid-template-columns: 1fr 1fr;
        align-items: start;
    }

    /* Remove mobile separators on desktop */
    .field:after {
        display: none;
    }

    /* Add subtle vertical divider between fields */
    .field:first-child:after {
        content: "";
        position: absolute;
        top: 0;
        right: -1rem;
        bottom: 0;
        width: 1px;
        background: linear-gradient(180deg, transparent, var(--border), transparent);
        opacity: 0.3;
        display: block;
    }
}

select {
    width: 100%;
    padding: clamp(0.75rem, 2vw, 1rem);
    font-size: clamp(1.5rem, 3vw, 1.875rem);  /* 24px-30px responsive - larger for VI */
    border: 2px solid var(--border);
    border-radius: 4px;  /* Subtle rounding for modern feel */
    background: var(--bg);
    color: var(--text);
    appearance: none;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath fill='%23505a5f' d='M1.5 4L6 8.5 10.5 4'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 1rem center;
    font-weight: 500;  /* Slightly bolder for better readability */
    line-height: 1.4;
}

select:focus {
    outline: 3px solid var(--focus);
    outline-offset: 0;
    box-shadow: inset 0 0 0 2px var(--focus-text);
    transform: scale(1.02);
    border-color: var(--primary);
}

/* Subtle animations for better UX */
select,
.file-list,
.form-grid .field {
    transition: all 0.2s ease;
}

select:hover {
    border-color: var(--primary);
    transform: translateY(-1px);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

.file-list {
    background: var(--bg-alt);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: clamp(1rem, 2vw, 1.5rem);
    margin-bottom: clamp(1.5rem, 3vw, 2rem);
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    transition: all 0.2s ease;
}

.file-list:hover {
    border-color: var(--primary);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

.file-list-title {
    font-weight: 700;
    margin-bottom: 1rem;
    color: var(--text-muted);
    font-size: clamp(1rem, 1.8vw, 1.125rem);  /* 16px-18px */
    text-transform: uppercase;
    letter-spacing: 0.08em;
    line-height: 1.2;
}

.file-list ul {
    margin: 0;
    padding-left: 1.5rem;
    line-height: 1.6;
}

.file-list li {
    padding: clamp(0.5rem, 1vw, 0.75rem) 0;
    font-family: ui-monospace, SFMono-Regular, "SF Mono", Monaco, "Consolas", monospace;
    font-size: clamp(1rem, 2vw, 1.125rem);
    font-weight: 500;
    color: var(--text);
    line-height: 1.4;
    border-bottom: 1px solid transparent;
    transition: all 0.15s ease;
}

.file-list li:hover {
    background: rgba(29, 112, 184, 0.05);
    border-bottom-color: var(--border);
    padding-left: 0.5rem;
    border-radius: 4px;
}

.file-list li:last-child {
    border-bottom: none;
}

/* File count indicator */
.file-count {
    display: inline-block;
    background: var(--primary);
    color: var(--primary-text);
    font-size: clamp(0.75rem, 1.5vw, 0.875rem);
    font-weight: 600;
    padding: 0.25rem 0.5rem;
    border-radius: 12px;
    margin-left: 0.5rem;
}

.actions {
    display: flex;
    flex-wrap: wrap;
    gap: clamp(0.75rem, 2vw, 1rem);
    margin-top: clamp(2rem, 4vw, 3rem);
    justify-content: center;
    padding: clamp(1rem, 2vw, 1.5rem);
    background: linear-gradient(135deg, var(--bg-alt), var(--bg));
    border-radius: 12px;
    border: 1px solid var(--border);
}

@media (min-width: 768px) {
    .actions {
        flex-direction: row;
        justify-content: center;
        gap: clamp(1rem, 3vw, 2rem);
    }

    .actions button {
        flex: 0 1 auto;
        min-width: 220px;
        max-width: 280px;
    }
}

/* Perfect balance for larger screens */
@media (min-width: 1024px) {
    .actions {
        gap: 3rem;
    }

    .actions button {
        min-width: 250px;
        max-width: 300px;
    }
}

button {
    padding: clamp(1rem, 2.5vw, 1.375rem) clamp(1.75rem, 4vw, 2.5rem);
    font-size: clamp(2rem, 5vw, 2.5rem);  /* 32px-40px responsive - HUGE for VI! */
    font-weight: 700;  /* Bolder weight for more impact */
    border: none;
    border-radius: 8px;  /* Slightly more rounded for modern feel */
    cursor: pointer;
    transition: all 0.2s ease;
    letter-spacing: 0.02em;  /* Slightly more spacing for larger text */
    line-height: 1.1;
    min-height: clamp(56px, 10vw, 64px);  /* Larger touch targets */
    text-transform: uppercase;  /* Make it more prominent */
}

button:focus {
    outline: 3px solid var(--focus);
    outline-offset: 0;
    box-shadow: inset 0 0 0 2px var(--focus-text);
}

.btn-primary {
    background: var(--primary);
    color: var(--primary-text);
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    position: relative;
    overflow: hidden;
}

.btn-primary:hover {
    background: var(--primary-hover);
    transform: translateY(-1px);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
}

.btn-primary:active {
    transform: translateY(0);
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.1);
}

.btn-secondary {
    background: var(--bg-alt);
    color: var(--text);
    border: 2px solid var(--border);
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    position: relative;
    overflow: hidden;
}

.btn-secondary:hover {
    background: var(--border);
    border-color: var(--primary);
    transform: translateY(-1px);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

.btn-secondary:active {
    transform: translateY(0);
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
}

/* Loading state for buttons */
button:disabled {
    opacity: 0.6;
    cursor: not-allowed;
    transform: none !important;
}

button:disabled:hover {
    transform: none;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

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
    font-size: clamp(1.5rem, 3.5vw, 1.875rem);  /* 24px-30px - match button size */
    margin-bottom: 1rem;
    line-height: 1.3;
    letter-spacing: -0.01em;
    display: flex;
    align-items: center;
    gap: 0.75rem;
}

.status-title::before {
    font-size: 1.5em;
    animation: statusPulse 0.6s ease-out;
}

@keyframes statusPulse {
    0% { transform: scale(0.8); opacity: 0; }
    50% { transform: scale(1.1); opacity: 1; }
    100% { transform: scale(1); opacity: 1; }
}

.log-output {
    font-family: ui-monospace, SFMono-Regular, "SF Mono", Monaco, "Consolas", monospace;
    font-size: clamp(0.9rem, 1.8vw, 1rem);  /* 14px-16px for code readability */
    line-height: 1.5;
    white-space: pre-wrap;
    max-height: 60vh;
    min-height: 200px;
    overflow-y: auto;
    background: var(--text);
    color: var(--bg);
    padding: clamp(1rem, 2vw, 1.5rem);
    margin-top: 1rem;
    border-radius: 6px;
    font-weight: 500;
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

/* Skip link for accessibility */
.skip-link {
    position: absolute;
    top: -40px;
    left: 8px;
    background: var(--primary);
    color: var(--primary-text);
    padding: 8px 16px;
    text-decoration: none;
    border-radius: 4px;
    font-weight: 600;
    font-size: 0.875rem;
    z-index: 1000;
    transition: all 0.2s ease;
}

.skip-link:focus {
    top: 8px;
    outline: 3px solid var(--focus);
    outline-offset: 2px;
}

/* Better focus indicators for form elements */
select:focus-visible,
button:focus-visible {
    outline: 3px solid var(--focus);
    outline-offset: 2px;
    box-shadow: inset 0 0 0 2px var(--focus-text);
}

/* Improved status announcements */
.status[role="alert"] {
    border-radius: 6px;
}

.status.success {
    border-color: var(--success);
    background: var(--success-bg);
}

.status.error {
    border-color: var(--error);
    background: var(--error-bg);
}

/* Progress bar styles */
.progress-container {
    margin: clamp(1.5rem, 3vw, 2rem) 0;
    padding: clamp(1rem, 2vw, 1.5rem);
    background: var(--bg-alt);
    border-radius: 8px;
    border: 1px solid var(--border);
}

.progress-bar {
    width: 100%;
    height: clamp(12px, 2vw, 16px);
    background: var(--border);
    border-radius: 8px;
    overflow: hidden;
    margin-bottom: 1rem;
    position: relative;
}

.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--primary), var(--primary-hover));
    border-radius: 8px;
    transition: width 0.3s ease;
    position: relative;
}

.progress-fill::after {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    right: 0;
    bottom: 0;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
    animation: progressShimmer 2s infinite;
}

@keyframes progressShimmer {
    0% { transform: translateX(-100%); }
    100% { transform: translateX(100%); }
}

.progress-stage {
    font-size: clamp(1.25rem, 2.5vw, 1.5rem);
    font-weight: 600;
    color: var(--primary);
    margin-bottom: 0.5rem;
}

.progress-message {
    font-size: clamp(1rem, 2vw, 1.25rem);
    color: var(--text-muted);
    font-style: italic;
}

/* Form disabled state */
.form-disabled {
    opacity: 0.6;
    pointer-events: none;
}
""")

app, rt = fast_app(hdrs=[CSS])

# Global progress tracking (simple and reliable)
pipeline_status = {
    "running": False,
    "progress": 0,
    "stage": "Ready",
    "message": "Ready to run pipeline",
    "logs": [],
    "success": None
}


# --- Helpers ---
def get_raw_files(service_type: str = None) -> list[Path]:
    if not RAW_DIR.exists():
        return []
    pattern = FILE_PATTERNS.get(service_type, "*.xlsx") if service_type else "*.xlsx"
    return sorted(RAW_DIR.glob(pattern), reverse=True)


def get_months(service_type: str) -> list[str]:
    """Extract month patterns (e.g., 'Aug-25') from filenames."""
    import re
    from fft.config import MONTH_ABBREV

    # Create pattern to match month-year format (e.g., Aug-25, Sep-24)
    month_abbrevs = '|'.join(MONTH_ABBREV.values())  # Jan|Feb|Mar|etc.
    pattern = rf'\b({month_abbrevs})-(\d{{2}})\b'

    months = set()
    for file_path in get_raw_files(service_type):
        # Look for month-year patterns in filename
        matches = re.findall(pattern, file_path.name)
        for month_abbrev, year in matches:
            months.add(f"{month_abbrev}-{year}")

    return sorted(months, reverse=True)


def update_progress(progress: int, stage: str, message: str):
    """Update the global progress state."""
    pipeline_status.update({
        "progress": progress,
        "stage": stage,
        "message": message
    })
    # Force immediate update
    logger.info(f"Progress updated: {progress}% - {stage}: {message}")

def run_cmd(service: str, month: str) -> tuple[bool, str]:
    """Run the pipeline command with progress tracking."""
    global pipeline_status

    # Completely reset and start progress tracking
    pipeline_status = {
        "running": True,
        "progress": 0,
        "stage": "Starting",
        "message": "Initializing pipeline...",
        "logs": [],
        "success": None
    }

    try:
        flag = [k for k, v in SERVICE_TYPES.items() if v == service][0]
        cmd = ["uv", "run", "python", "-m", "fft", f"--{flag}"]
        if month and month != "all":
            cmd.extend(["--month", month])

        project_root = Path(__file__).parent.parent
        logger.info(f"Running: {' '.join(cmd)}")

        # Progress stages
        update_progress(10, "Loading", f"Loading {service} data files...")

        update_progress(25, "Processing", f"Processing {service} data...")

        # Run the actual command
        update_progress(50, "Running", "Executing FFT pipeline...")
        result = subprocess.run(cmd, capture_output=True, text=True, cwd=project_root)

        update_progress(75, "Finishing", "Finalizing output...")

        success = result.returncode == 0
        output = result.stdout if success else (result.stderr or result.stdout)

        # Store logs for final display
        pipeline_status["logs"] = [output] if output else ["No output captured."]

        # Complete
        if success:
            update_progress(100, "Complete", "Pipeline completed successfully!")
        else:
            update_progress(100, "Failed", "Pipeline execution failed")

        pipeline_status.update({
            "running": False,
            "success": success
        })

        logger.info(f"Return code: {result.returncode}")
        logger.info(f"Pipeline status after completion: running={pipeline_status['running']}, success={pipeline_status['success']}")
        return success, output

    except Exception as e:
        pipeline_status.update({
            "running": False,
            "progress": 100,
            "stage": "Error",
            "message": f"Error: {str(e)}",
            "success": False
        })
        return False, f"Error running pipeline: {str(e)}"


# --- Progress Components ---
def progress_bar(progress: int):
    """Create a progress bar component."""
    return Div(
        Div(
            style=f"width: {progress}%",
            cls="progress-fill"
        ),
        cls="progress-bar"
    )

def progress_display():
    """Create the complete progress display."""
    # Always return a div, but show/hide content based on state
    if not pipeline_status["running"]:
        # When not running, return empty div and re-enable form
        return Div(
            Script("""
                document.getElementById('main-form').classList.remove('form-disabled');
                var submitBtn = document.querySelector('[type="submit"]');
                if (submitBtn) submitBtn.disabled = false;
            """),
            style="display: none;"  # Hidden when not running
        )

    # Show progress bar and status only while running
    content = [
        progress_bar(pipeline_status["progress"]),
        Div(pipeline_status["stage"], cls="progress-stage"),
        Div(pipeline_status["message"], cls="progress-message"),
        Script("""
            document.getElementById('main-form').classList.add('form-disabled');
            var submitBtn = document.querySelector('[type="submit"]');
            if (submitBtn) submitBtn.disabled = true;
        """)
    ]

    return Div(*content, cls="progress-container")

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
        aria_label="Select NHS service type for data processing",
        aria_describedby="service-help",
    )


def month_select(months=None):
    opts = [Option("All months", value="all", selected=True)]
    if months:
        opts += [Option(m, value=m) for m in months]
    return Select(
        *opts,
        name="month",
        id="month",
        aria_label="Select month for data processing",
        aria_describedby="month-help"
    )


def file_list_box(files):
    if not files:
        return Div(
            Div("Available files", cls="file-list-title"),
            P("No files found in raw folder",
              style="padding: 1rem 0; color: var(--text-muted); font-style: italic;"),
            cls="file-list",
            role="region",
            aria_label="Available data files",
        )

    # Show more files and add count indicator
    display_files = files[:12]  # Increased from 8 to 12
    items = [Li(f.name, title=f"File: {f.name}") for f in display_files]

    if len(files) > 12:
        extra = Li(
            f"... and {len(files) - 12} more files",
            style="font-style: italic; color: var(--text-muted); opacity: 0.8;"
        )
        items.append(extra)

    title_with_count = Div(
        "Available files",
        Span(str(len(files)), cls="file-count", title=f"{len(files)} total files"),
        cls="file-list-title"
    )

    return Div(
        title_with_count,
        Ul(*items),
        cls="file-list",
        role="region",
        aria_label=f"Available data files ({len(files)} total)",
    )


def status_box(success: bool, msg: str, log: str = None):
    cls = "status success" if success else "status error"
    icon = "✓" if success else "✗"

    # Create title with separate icon element for better styling
    title = Div(
        Span(icon, style="font-size: 1.5em;"),
        Span(msg),
        cls="status-title"
    )

    content = [title]
    if log:
        content.append(
            Details(
                Summary("View detailed log output", style="font-size: 1.1rem; margin-top: 1rem;"),
                Div(log, cls="log-output", role="log", aria_live="polite"),
            )
        )
    return Div(*content, cls=cls, role="alert", aria_live="polite")


# --- Routes ---
@rt("/")
def get():
    return Titled(
        "FFT Pipeline",
        # Skip link for accessibility
        A("Skip to main content", href="#main-form", cls="skip-link"),
        # Progress display area (polls every 2 seconds when pipeline is running)
        Div(
            id="progress-area",
            hx_get="/progress",
            hx_trigger="load, every 2s",
            hx_swap="innerHTML"
        ),
        Form(
            Div(
                Div(
                    Label("Service Type", for_="service"),
                    service_select(),
                    Div("Choose the NHS service type to process", id="service-help", cls="visually-hidden"),
                    cls="field"
                ),
                Div(
                    Label("Month", for_="month"),
                    Div(month_select(), id="month-container"),
                    Div("Select a specific month or process all available data", id="month-help", cls="visually-hidden"),
                    cls="field",
                ),
                cls="form-grid",
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
            id="main-form",
            role="main",
            aria_label="FFT Pipeline Configuration and Execution",
            hx_post="/run",
            hx_target="#status",
        ),
    )


@rt("/months")
def get(service: str = ""):
    return month_select(get_months(service))


@rt("/progress")
def get_progress():
    """Return current progress status."""
    return progress_display()


@rt("/status-check")
def get_status_check():
    """Check if pipeline is complete and return final status."""
    # Debug info
    logger.info(f"Status check: running={pipeline_status['running']}, success={pipeline_status['success']}, stage={pipeline_status['stage']}, progress={pipeline_status['progress']}")

    try:
        if pipeline_status["running"]:
            # Still running, keep checking
            return Div(
                f"Pipeline running... (Stage: {pipeline_status['stage']}, {pipeline_status['progress']}%)",
                style="padding: 1rem; background: var(--bg-alt); border-radius: 6px; margin-top: 1rem;",
                hx_get="/status-check",
                hx_trigger="every 2s",
                hx_swap="innerHTML",
                id="pipeline-status"
            )
        elif pipeline_status["success"] is not None:
            # Pipeline complete, show final result
            success = pipeline_status["success"]
            # Get the logs from the global state if available
            log_output = "\n".join(pipeline_status["logs"]) if pipeline_status["logs"] else "Pipeline execution completed."

            logger.info(f"Pipeline completed with success={success}, showing final result")

            # Don't clear status immediately - let it persist for display
            # It will be reset when a new pipeline starts

            return status_box(
                success,
                "Pipeline completed successfully" if success else "Pipeline failed",
                log_output
            )
        else:
            # No result yet, keep waiting
            return Div(
                "Waiting for pipeline to start...",
                style="padding: 1rem; background: var(--bg-alt); border-radius: 6px; margin-top: 1rem;",
                hx_get="/status-check",
                hx_trigger="every 2s",
                hx_swap="innerHTML",
                id="pipeline-status"
            )
    except Exception as e:
        logger.error(f"Error in status check: {e}")
        return Div(
            f"Error checking status: {str(e)}",
            style="padding: 1rem; background: var(--error-bg); color: var(--error); border-radius: 6px; margin-top: 1rem;",
            hx_get="/status-check",
            hx_trigger="every 2s",
            hx_swap="innerHTML",
            id="pipeline-status"
        )


@rt("/files")
def get(service: str = ""):
    return file_list_box(get_raw_files(service or None))


@rt("/run")
async def post(service: str, month: str):
    if not service:
        return status_box(False, "Please select a service type")

    # Start the pipeline asynchronously so progress can be seen
    import asyncio
    import threading

    def run_pipeline_thread():
        try:
            run_cmd(service, month)
        except Exception as e:
            # Ensure pipeline status is reset even if there's an exception
            pipeline_status.update({
                "running": False,
                "progress": 100,
                "stage": "Error",
                "message": f"Error: {str(e)}",
                "success": False,
                "logs": [f"Pipeline error: {str(e)}"]
            })

    # Start pipeline in background thread
    thread = threading.Thread(target=run_pipeline_thread)
    thread.daemon = True  # Allow main thread to exit
    thread.start()

    # Return immediate response to show progress has started
    return Div(
        "Pipeline started!",
        style="padding: 1rem; background: var(--bg-alt); border-radius: 6px; margin-top: 1rem;",
        hx_get="/status-check",
        hx_trigger="every 2s",
        hx_swap="innerHTML",
        id="pipeline-status"
    )


@rt("/open-output")
def post():
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    webbrowser.open(f"file://{OUTPUTS_DIR.absolute()}")
    return ""


def cleanup_port_5001():
    """Kill any processes using port 5001 to ensure clean startup."""
    import platform
    import time
    import re

    try:
        if platform.system() == "Windows":
            # Windows: use netstat and taskkill
            result = subprocess.run(
                ["netstat", "-ano"],
                capture_output=True,
                text=True
            )
            if result.returncode == 0:
                pids = []
                for line in result.stdout.split('\n'):
                    if ':5001' in line and 'LISTENING' in line:
                        # Extract PID from last column
                        parts = line.split()
                        if parts:
                            pid = parts[-1]
                            if pid.isdigit():
                                pids.append(pid)

                for pid in pids:
                    subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True)

                if pids:
                    logger.info(f"Cleaned up processes on port 5001: {pids}")
                    time.sleep(0.5)
        else:
            # Unix/Linux/macOS: use lsof and kill
            result = subprocess.run(["lsof", "-ti:5001"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                pids = result.stdout.strip().split()
                for pid in pids:
                    subprocess.run(["kill", pid], capture_output=True)
                logger.info(f"Cleaned up processes on port 5001: {pids}")
                time.sleep(0.5)
    except Exception as e:
        logger.debug(f"Port cleanup failed (likely no processes to clean): {e}")


if __name__ == "__main__":
    cleanup_port_5001()
    serve()
