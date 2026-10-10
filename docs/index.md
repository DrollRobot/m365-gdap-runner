# m365-gdap-runner

Run one Python payload across every customer tenant you reach through GDAP, from a single
partner sign-in. The runner owns the loop: sign-in, tenant listing, consent checks, error
isolation, parallelism and merged output. Authentication and the service clients come from
[azure-auth](https://github.com/DrollRobot/azure-auth). A payload contains neither.

## Installation

```bash
uv tool install git+https://github.com/DrollRobot/m365_gdap_runner.git
m365-gdap-runner-config init      # asks for the partner tenant and your sign-in name
```

## Quick start

A payload is a `.py` file with `run(ctx)` and two declarations:

```python
from m365_gdap_runner import TenantContext

CLIENTS = ["graph"]  # which of ctx.graph, ctx.exchange, ctx.ipps, ctx.azure you use
SCOPES = ["User.Read.All"]  # delegated Graph scopes; only when "graph" is in CLIENTS


def run(ctx: TenantContext) -> list[dict[str, str]]:
    users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
    return [{"userPrincipalName": u["userPrincipalName"]} for u in users]


if __name__ == "__main__":
    import json
    from m365_gdap_runner import standalone

    with standalone(SCOPES, CLIENTS) as ctx:
        print(json.dumps(run(ctx), indent=2))
```

Run it across every tenant, or against one:

```bash
m365-gdap-runner users.py --jobs 5 --format csv --output users.csv
python users.py --tenant Contoso
```

The runner signs in once per client in `CLIENTS` at the partner tenant (each uses a
different Microsoft app), then for every customer tenant checks each client silently before
calling `run()`. A tenant that has not consented, or wants a fresh sign-in, is reported as
`auth_failed` and skipped. With `--cache disk` later runs skip the browser.

`ctx.auth` is the azure-auth context for the tenant. Hand it to any azure-auth client, or to
an Azure SDK client as `credential=`, for resources the four properties do not cover.

See `examples/` for a Graph payload and an Exchange Online payload.
