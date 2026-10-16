# m365-gdap-runner

Run a Python script across every customer tenant you reach through GDAP, from a single
partner sign-in. The library owns the loop: sign-in, tenant listing, consent checks, error
isolation, parallelism and merged output. Authentication and the service clients come from
[azure-auth](https://github.com/DrollRobot/azure-auth). A script contains neither.

## Installation

```bash
uv tool install git+https://github.com/DrollRobot/m365-gdap-runner.git
m365-gdap-runner-config init      # asks for the partner tenant and your sign-in name
m365-gdap-runner                  # creates tenants.csv; add TRUE/FALSE group columns
```

Scripts do not need this install: each one names the library in its header, and `uv run`
builds the script's own environment from that.

## Quick start

A script is a `.py` file with `run(ctx)`, two declarations, and a `__main__` block that hands
it to the library. The header at the top tells uv what the script needs:

```python
# /// script
# requires-python = ">=3.14"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
from m365_gdap_runner import TenantContext

CLIENTS = ["graph"]  # which of ctx.graph, ctx.exchange, ctx.ipps, ctx.azure you use
SCOPES = ["User.Read.All"]  # delegated Graph scopes; only when "graph" is in CLIENTS


def run(ctx: TenantContext) -> list[dict[str, str]]:
    users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
    return [{"userPrincipalName": u["userPrincipalName"]} for u in users]


if __name__ == "__main__":
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
```

Run it across a group of tenants, or against one:

```bash
uv run users.py --group pilot --jobs 5 --format csv --output users.csv
uv run users.py --tenant Contoso
```

The runner signs in once per client in `CLIENTS` at the partner tenant (each uses a
different Microsoft app), then for every selected tenant checks each client silently before
calling `run()`. A tenant that has not consented, or wants a fresh sign-in, is reported as
`auth_failed` and skipped. With `--cache disk` later runs skip the browser.

## Flags of your own

`main(__name__)` reads the library's flags only. A script with flags of its own builds the
parser, adds the library's with `add_arguments()`, and passes what it parsed to `main()`;
`-h` then lists both:

```python
if __name__ == "__main__":
    import argparse
    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(description="Licensed users.")
    parser.add_argument("--sku", help="Only users with this SKU.")
    add_arguments(parser)
    args = parser.parse_args()
    SKU = args.sku
    raise SystemExit(main(__name__, args))
```

`run()` reads `SKU` like any other module-level value.

## Choosing tenants

A run has to say which tenants it is for; one that names none is refused. Pick them with:

- `--group NAME`: the tenants marked TRUE in column `NAME` of the tenants file. Repeat it
  to add groups together.
- `--tenant NAME_OR_ID`: one tenant by display name or ID. Repeatable, and combines with
  `--group`.
- `--all`: every tenant with a GDAP relationship.

`--exclude NAME_OR_ID` drops tenants from any of these.

The tenants file is a CSV you edit in a spreadsheet, one row per customer. This package owns
four columns, `m365_tenant_id`, `m365_tenant_name`, `m365_domain` and `m365_cloud`, and
reads one more per `--group`: a column holding TRUE or FALSE, where a blank cell counts as
FALSE and anything else is an error. Every other column is left alone, so the same file can
carry what other tools know about the same customers on other platforms. Tenants are matched
by ID; the name and domain columns are there for you. A row with no `m365_tenant_id` is a
customer without an M365 tenant and is skipped.

```csv
m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,pilot,patching,cw_company_id
aaaaaaaa-0000-0000-0000-00000000000a,Alpha,alpha.example,Commercial,TRUE,FALSE,1001
bbbbbbbb-0000-0000-0000-00000000000b,beta corp,beta.example,USGov,TRUE,TRUE,1002
```

`m365_cloud` is `Commercial`, `USGov`, `USGovDoD` or `China`. The file tracks tenants in
every cloud, but a run signs in to one, so a group member in another cloud is skipped and the
log says so. The runner fills the cell in itself, by OpenID Connect discovery, when it appends
a tenant or finds the cell blank.

The file is `tenants.csv` in the config directory (`m365-gdap-runner-config path` shows
where), unless the `tenants_file` setting or `--tenants-file` points elsewhere. Every run
keeps it current: a GDAP tenant with no row is appended with its name, its cloud and every
other column blank, and a warning names it. The `m365-gdap-runner` command does that, lists
the tenants and runs nothing, so run it once to create the file, then add group columns. A row whose tenant
no longer has a GDAP relationship is left alone; selecting it with `--group` is an error.

`ctx.auth` is the azure-auth context for the tenant. Hand it to any azure-auth client, or to
an Azure SDK client as `credential=`, for resources the four properties do not cover.

See `examples/` for a Graph script and two Exchange Online scripts. In this repository,
`uv run python examples/licensed_users.py ...` runs one against the working copy instead of
the release its header names.
