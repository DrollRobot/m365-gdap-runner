# Development

Setting up a working copy, to change the package itself.

## Requirements

- [uv](https://docs.astral.sh/uv/). It fetches the Python version pinned in
  `.python-version`, 3.11, when the system has none.
- git.

## Set up

```bash
git clone https://github.com/DrollRobot/m365-gdap-runner.git
cd m365-gdap-runner
uv sync
uv run pre-commit install
uv run pre-commit install --hook-type pre-push
```

## Configuration

The working copy reads the same config file and tenants file as an installed copy. To keep
others for development, point `M365_GDAP_RUNNER_CONFIG_DIR` at another directory, and run
`m365-gdap-runner-config init` there:

```bash
export M365_GDAP_RUNNER_CONFIG_DIR=~/m365-gdap-runner-dev
uv run m365-gdap-runner-config init
```

## Running a script against the working copy

A script's header names a release of the runner, so `uv run script.py` runs that release.
To run a script against the working copy instead, run it with the project's Python:

```bash
uv run python path/to/script.py --include pilot
```

Anything else the script's header names must then be in the project's environment too.

## Tests

See [AGENTS.TESTING.md](https://github.com/DrollRobot/m365-gdap-runner/blob/main/AGENTS.TESTING.md)
in the repository root.

## Docs

```bash
uv run mkdocs serve           # live preview at http://127.0.0.1:8000
uv run mkdocs build --strict  # build, failing on any warning
```
