# Getting started

## Requirements

- [uv](https://docs.astral.sh/uv/). It fetches Python 3.11 or newer when the system has
  none.
- A partner account with GDAP access to the customer tenants you want to run scripts
  against.

## Install

The package has two commands: `m365-gdap-runner-config` sets up the configuration, and
`m365-gdap-runner-customer-export` creates the tenants file. Install them with:

```bash
uv tool install git+https://github.com/DrollRobot/m365-gdap-runner.git
```

Scripts do not need this install: each one names the library in its header, and `uv run`
builds the script's own environment from that.

## Configure

```bash
m365-gdap-runner-config init
```

It asks for each setting in turn. Two are required: `partner_tenant_id`, the partner
tenant's ID or domain name, and `username`, the partner account you sign in with. Press Enter
at the others to keep their defaults; [Settings](usage.md#settings) lists them all.

## Create the tenants file

```bash
m365-gdap-runner-customer-export
```

It signs in as the partner account, in the browser, lists every tenant with a GDAP
relationship, and writes them to `tenants.csv` in the config directory
(`m365-gdap-runner-config path` shows where).

Open the file in a spreadsheet and add a group column: a header such as `pilot`, and TRUE in
the rows of two or three tenants to try a first script on. A blank cell counts as FALSE.

## Write a script

Save this as `users.py`:

```python
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
from m365_gdap_runner import TenantContext, TenantResult

CLIENTS = ["graph"]  # which of ctx.graph, ctx.exchange, ctx.ipps, ctx.arm you use
SCOPES = ["User.Read.All"]  # delegated Graph scopes; only when "graph" is in CLIENTS


def run(ctx: TenantContext) -> list[str]:
    users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
    return [u["userPrincipalName"] for u in users]


def report(results: list[TenantResult]) -> None:
    for result in results:
        print(result.tenant.display_name, result.status, len(result.result or []))


if __name__ == "__main__":
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
```

`run(ctx)` does the work for one tenant. `report(results)` is optional; it gets every
tenant's result once the last one is done.

## Run it

```bash
uv run users.py --include pilot
```

The runner signs in at the partner tenant, in the browser, checks each tenant in the `pilot`
group through that sign-in, and calls `run()` for each one. Logs and a summary go to stderr;
what `report()` prints goes to stdout.

`--include` also takes a single tenant, by display name or ID, or `all` for every GDAP
tenant:

```bash
uv run users.py --include Contoso
```

`uv run users.py -h` lists every flag. [Usage](usage.md) covers the rest.
