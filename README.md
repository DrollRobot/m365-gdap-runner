# m365-gdap-runner

<!-- FIXME: point these badges at your own repo, and drop the ones you don't use -->
[![CI](https://github.com/DrollRobot/m365_gdap_runner/actions/workflows/ci.yml/badge.svg)](https://github.com/DrollRobot/m365_gdap_runner/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.14%2B-blue)](https://www.python.org/)
<!-- FIXME: match the badge and the link below to the license you chose -->
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Run one Python payload across every customer tenant you reach through GDAP, from a single
partner sign-in. The runner owns the loop: sign-in, tenant listing, consent checks, error
isolation, parallelism and merged JSON/CSV output. Authentication and the service clients
come from [azure-auth](https://github.com/DrollRobot/azure-auth); payloads contain neither.

## Installation

<!-- FIXME: replace with the real install command once the package is published -->
```bash
uv add git+https://github.com/DrollRobot/m365_gdap_runner.git
```

## Usage

A payload declares which clients it uses and, for Graph, which scopes, then implements
`run(ctx)`:

```python
from m365_gdap_runner import TenantContext

CLIENTS = ["graph"]
SCOPES = ["User.Read.All"]


def run(ctx: TenantContext) -> list[dict[str, str]]:
    users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
    return [{"userPrincipalName": u["userPrincipalName"]} for u in users]
```

```bash
m365-gdap-runner-config init                         # partner tenant and sign-in name
m365-gdap-runner --list-tenants                      # creates tenants.csv; add TRUE/FALSE group columns
m365-gdap-runner users.py --group pilot --format csv --output users.csv
```

Tenants are chosen with `--group` (a column of the tenants CSV), `--tenant` or `--all`, less
`--exclude`. `ctx` also offers `exchange`, `ipps` and `azure` clients when declared in
`CLIENTS`, and `ctx.auth`, the azure-auth context, for anything else. See
[docs/index.md](docs/index.md) and `examples/`.

## Development

```bash
uv sync                                     # create the venv, install dev/test deps
uv run pre-commit install                   # commit hooks
uv run pre-commit install --hook-type pre-push
uv run pytest -m "not live"                 # offline tests
```

<!-- FIXME: anything a contributor needs that the commands above don't cover --
     required config options, credentials, external services, test fixtures -->

## License

<!-- FIXME: name the license you chose; the full text is in LICENSE -->
