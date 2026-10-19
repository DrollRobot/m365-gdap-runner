# m365-gdap-runner

Run a Python script across every customer tenant you reach through GDAP, from a single
partner sign-in. The library owns the loop: sign-in, tenant listing, consent checks, error
isolation and parallelism. Authentication and the service clients come from
[azure-auth](https://github.com/DrollRobot/azure-auth). A script contains neither; it holds
the work for one tenant and its own reports.

## Installation

```bash
uv tool install git+https://github.com/DrollRobot/m365-gdap-runner.git
m365-gdap-runner-config init      # asks for the partner tenant and your sign-in name
m365-gdap-runner                  # creates tenants.csv; add TRUE/FALSE group columns
```

Scripts do not need this install: each one names the library in its header, and `uv run`
builds the script's own environment from that.

## Quick start

A script is a `.py` file with `run(ctx)`, two declarations, an optional `report(results)`,
and a `__main__` block that hands it to the library. The header at the top tells uv what the
script needs:

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

Run it across a group of tenants, or against one:

```bash
uv run users.py --include pilot --jobs 5
uv run users.py --include Contoso
```

The runner signs in once per client in `CLIENTS` at the partner tenant (each uses a
different Microsoft app), then, before the first `run()`, checks every selected tenant
silently through that sign-in. A tenant it cannot reach that way gets a sign-in of its own
(see [Tenants outside GDAP](#tenants-outside-gdap)).

The check asks for exactly the scopes in `SCOPES`, so a tenant that has not granted one of
them fails before `run()`. `SCOPES = []` asks for none in particular: Graph gets whatever
the app holds in the tenant, and only a tenant where it holds nothing fails. That suits a
script that reads what the app holds, with `ctx.graph.granted_scopes()`.

Sign-ins open the system's default browser. The Windows sign-in window (the broker) is not
offered: GDAP does not work through it. To use another browser, or another profile of it:

- `browser_path`: the browser, as a full path (recommended) or a program on `PATH`.
- `browser_args`: a list of arguments for it, such as `["--profile-directory=Profile 2"]`.
  The sign-in address is added last. Needs `browser_path`.

## Keeping sign-ins between runs

Sign-ins are kept in memory and end with the run. To keep them, opt in to the disk cache:
the `cache` setting, or `--cache`, set to `disk` keeps them in an encrypted file, and later
runs skip the sign-in. That includes each tenant's own sign-in: a tenant the partner sign-in
cannot reach uses its cached sign-in, and opens the browser only when there is none.

The file holds refresh tokens, which sign in without a password. Two settings help keep
track of it:

- `cache_path`: the file to keep it in. Empty uses azure-auth's file in the per-user cache
  directory.
- `cache_reminder_path`: a folder in plain sight, such as the Desktop. Every run with the
  disk cache puts `M365-GDAP-RUNNER_MSAL_CACHE_DELETEME` there, a link to the folder that
  holds the cache: a `.url` shortcut on Windows, a symbolic link elsewhere. It holds only a
  path, so it is safe in a synced folder; the cache itself is not.

## Reports

The library writes no output of its own; reports belong to the script, in whatever form it
needs.

- **Per tenant:** `run(ctx)` writes it, with `ctx.tenant` for the tenant's name and ID.
- **Per run:** `report(results)`, when the script has one, is called once after the last
  tenant. It gets one `TenantResult` per tenant in run order, failed ones included: `tenant`,
  `status` (`ok`, `auth_failed`, `error` or `skipped`), `result` (whatever `run()` returned,
  as it is) and `error`.

`report()` is not called after Ctrl+C, since a report of part of a run would look complete.
If it raises, the error is logged and the run exits with code 2. Logs and the summary go to
stderr, so stdout is the script's. `examples/mailbox_audit_gaps.py` writes an Excel workbook
of each kind.

## Errors in run()

An error raised in `run()` stops the run: that tenant is `error`, and the tenants not yet
started are `skipped`. That keeps a script that changes things from repeating a broken
change in every tenant. A script that only reads sets `CONTINUE_ON_ERROR = True` to keep
going. `--continue-on-error` and `--no-continue-on-error` override the script.

A tenant that could not be reached is `auth_failed` and never stops the run.

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

A run has to say which tenants it is for; one that includes none is refused. Pick them with
`--include`, and take some back out with `--exclude`. Both are repeatable, and each takes:

- a tenant, by display name or ID: a GDAP customer or a row of the tenants file.
- a group: the tenants marked TRUE in that column of the tenants file.
- `all`, for `--include` only: every tenant with a GDAP relationship.

Everything included adds up, and an exclusion always wins: `--include all --exclude pilot`
runs every GDAP tenant outside the pilot group. Case does not matter. A name that matches
nothing is an error, and so is one that could mean two things: a tenant and a group column
with the same name, or a tenant or column named `all`. Name the tenant by its ID, or rename
the column. `--include m365_gdap` runs the tenants the partner sign-in reaches (see
[Tenants outside GDAP](#tenants-outside-gdap)).

The tenants file is a CSV you edit in a spreadsheet, one row per customer. This package owns
six columns, `m365_tenant_id`, `m365_tenant_name`, `m365_domain`, `m365_cloud`,
`m365_username` and `m365_gdap`, and reads one more per group a run names:
a column holding TRUE or FALSE, where a blank cell counts as FALSE and anything else is an
error. Every other column is left alone, so the same file can carry what other tools know
about the same customers on other platforms. Tenants are matched by ID; the name and domain
columns are there for you. A row with no `m365_tenant_id` is a customer without an M365
tenant and is skipped.

```csv
m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,m365_gdap,pilot,patching,cw_company_id
aaaaaaaa-0000-0000-0000-00000000000a,Alpha,alpha.example,Commercial,,TRUE,TRUE,FALSE,1001
bbbbbbbb-0000-0000-0000-00000000000b,beta corp,beta.example,USGov,admin@beta.us,,TRUE,TRUE,1002
cccccccc-0000-0000-0000-00000000000c,Gamma,gamma.example,Commercial,admin@gamma.example,FALSE,FALSE,TRUE,1003
```

`m365_cloud` is `Commercial`, `USGov`, `USGovDoD` or `China`. The runner fills the cell in
itself, by OpenID Connect discovery, when it appends a tenant or finds the cell blank.

The file is `tenants.csv` in the config directory (`m365-gdap-runner-config path` shows
where), unless the `tenants_file` setting or `--tenants-file` points elsewhere. Every run
keeps it current. Any of the six `m365_` columns the file lacks is added after the `m365_`
columns it has (or at the end when it has none), blank, except that `m365_gdap` is TRUE for
every tenant GDAP lists. A GDAP tenant with no row is appended with its name, its cloud, TRUE
in `m365_gdap` and every other column blank, and a warning names it. The `m365-gdap-runner` command does that, lists
the tenants and runs nothing, so run it once to create the file, then add group columns. Rows
for tenants without a GDAP relationship are yours to add and keep.

## Tenants outside GDAP

A tenant the partner sign-in cannot reach still runs. That covers a row of the tenants file
with no GDAP relationship, a tenant in another cloud (the partner sign-in reaches its own
cloud only), and a GDAP tenant where consent is missing or a policy wants a fresh sign-in.

Before the first `run()`, every selected tenant in the partner's cloud is checked silently
through the partner sign-in; every 30 seconds, a log line says how many are done and about how
long the rest will take. Each tenant that fails that check, and each tenant in another
cloud, then gets a sign-in of its own, in its own cloud, as the account in its
`m365_username` cell. That is one sign-in per client in `CLIENTS`, one tenant at a
time, all before the loop starts, so `--jobs` runs never prompt. With `--cache disk`, later
runs are silent. A sign-in that fails or is cancelled leaves that tenant `auth_failed`; the
rest run.

A tenant whose `m365_username` cell is blank gets no sign-in of its own: it is `auth_failed`,
with the reason, and no sign-in window opens for it.

The `m365_gdap` column says which tenants the partner sign-in reaches. TRUE, which every
appended row gets, tries the partner sign-in first. FALSE or blank never tries it: the tenant
goes straight to a sign-in of its own. GDAP can list a customer whose tenant it gives you no
access to; put FALSE in that row's cell. With `m365_username` blank as well, the tenant is
left out of the run, with a warning, instead of failing: that marks a customer you cannot
reach at all but cannot remove from Partner Center. Anything else in the cell is an error.

`m365_gdap` is a group like any other, so `--include m365_gdap` runs only the tenants marked
TRUE. A tenant there whose partner sign-in fails still gets a sign-in of its own when its
`m365_username` is filled in.

`ctx.auth` is the azure-auth context for the tenant. Hand it to any azure-auth client, or to
an Azure SDK client as `credential=`, for resources the four properties do not cover.

See `examples/` for two Graph scripts and three Exchange Online scripts. In this repository,
`uv run python examples/licensed_users.py ...` runs one against the working copy instead of
the release its header names.
