# Contributing

Thanks for your interest in contributing! This guide will help you get started.

## Quick Start

1. **Fork** this repo (top-right button on GitHub)
2. **Clone** your fork: `git clone https://github.com/YOUR_USERNAME/fft_pipeline.git`
3. **Add upstream** remote: `git remote add upstream https://github.com/nhsengland/fft_pipeline.git`
4. **Create a branch**: `git checkout -b my-feature`
5. **Make changes**, commit with clear messages
6. **Push** to your fork: `git push origin my-feature`
7. **Open a Pull Request** from your fork to `main`

## Why Fork?

You likely don't have push access to this repo. Forking creates your own copy where you can push freely, then propose changes via Pull Request.

## Before You Start

- Check existing [Issues](../../issues) and [Pull Requests](../../pulls) to avoid duplicates
- For large changes, open an issue first to discuss

## Guidelines

- Keep PRs focused — one feature or fix per PR
- Write clear commit messages: `Add X` / `Fix Y` / `Update Z`
- Test your changes locally before submitting
- Update documentation if needed

## Staying Up to Date

```bash
git fetch upstream
git checkout main
git merge upstream/main
git push origin main
```

Then rebase your feature branch if needed: `git rebase main`

## Code Style

This project follows the **Google Python Style Guide** with **Ruff** for linting/formatting and **Ty** for type checking:

- **Style guide**: [Google Python Style Guide](https://google.github.io/styleguide/pyguide.html)
- **Line length**: 90 characters
- **Python version**: 3.13+
- **Quote style**: Double quotes
- **Indentation**: 4 spaces (no tabs)
- **Docstrings**: Google-style format

### Before committing, run:

```bash
# Format code
uv run ruff format

# Check linting
uv run ruff check

# Type check
uv run ty

# Run doctests
uv run python -m doctest $(find src/fft/ -name "*.py" -not -name "__main__.py")
```

### Testing

- **Primary testing**: Doctests (inline with functions)
- **No pytest** - use doctests for documentation + testing
- **Validation**: `uv run python -m fft --validate`

## Questions?

Open an issue or start a discussion. We're happy to help!
