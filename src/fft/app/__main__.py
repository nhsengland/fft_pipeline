"""Entry point for the FFT Pipeline FastHTML web application.

Run with: uv run python -m fft.app
"""

if __name__ == "__main__":
    from fft.app.server import cleanup_port_5001

    cleanup_port_5001()

    # Run the server with proper import string for reload functionality
    import uvicorn

    uvicorn.run("fft.app.server:app", host="0.0.0.0", port=5001, reload=True)
else:
    # For uvicorn auto-discovery
    pass
