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

Run it across a group of tenants, or against one:

```bash
m365-gdap-runner users.py --group pilot --jobs 5 --format csv --output users.csv
python users.py --tenant Contoso
```

The runner signs in once per client in `CLIENTS` at the partner tenant (each uses a
different Microsoft app), then for every selected tenant checks each client silently before
calling `run()`. A tenant that has not consented, or wants a fresh sign-in, is reported as
`auth_failed` and skipped. With `--cache disk` later runs skip the browser.

## Choosing tenants

A run has to say which tenants it is for; one that names none is refused. Pick them with:

- `--group NAME`: the tenants marked TRUE in column `NAME` of the tenants file. Repeat it
  to add groups together.
- `--tenant NAME_OR_ID`: one tenant by display name or ID. Repeatable, and combines with
  `--group`.
- `--all`: every tenant with a GDAP relationship.

`--exclude NAME_OR_ID` drops tenants from any of these.

The tenants file is a CSV you edit in a spreadsheet, one row per customer. This package owns
three columns, `m365_tenant_id`, `m365_tenant_name` and `m365_domain`, and reads one more
per `--group`: a column holding TRUE or FALSE, where a blank cell counts as FALSE and
anything else is an error. Every other column is left alone, so the same file can carry what
other tools know about the same customers on other platforms. Tenants are matched by ID; the
name and domain columns are there for you. A row with no `m365_tenant_id` is a customer
without an M365 tenant and is skipped.

```csv
m365_tenant_id,m365_tenant_name,m365_domain,pilot,patching,cw_company_id
aaaaaaaa-0000-0000-0000-00000000000a,Alpha,alpha.example,TRUE,FALSE,1001
bbbbbbbb-0000-0000-0000-00000000000b,beta corp,beta.example,TRUE,TRUE,1002
```

The file is `tenants.csv` in the config directory (`m365-gdap-runner-config path` shows
where), unless the `tenants_file` setting or `--tenants-file` points elsewhere. Every run
keeps it current: a GDAP tenant with no row is appended with its name and every other column
blank, and a warning names it. `m365-gdap-runner --list-tenants` does that and nothing else,
so run it once to create the file, then add group columns. A row whose tenant no longer has
a GDAP relationship is left alone; selecting it with `--group` is an error.

`ctx.auth` is the azure-auth context for the tenant. Hand it to any azure-auth client, or to
an Azure SDK client as `credential=`, for resources the four properties do not cover.

See `examples/` for a Graph payload and an Exchange Online payload.
