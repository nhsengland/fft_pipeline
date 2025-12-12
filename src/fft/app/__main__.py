"""Entry point for the FFT Pipeline FastHTML web application.

Run with: uv run python -m fft.app
"""

# Import app at module level for FastHTML auto-discovery
from .server import app, serve

# Make app available in global namespace
__all__ = ["app"]

if __name__ == "__main__":
    serve()