# Usage

## Scripts

A script is a `.py` file that does the work for one tenant; the runner calls it once per
selected tenant. It defines `run(ctx)`, declares what it uses, and hands itself to the
library from its `__main__` block:

```python
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "m365-gdap-runner @ git+https://github.com/DrollRobot/m365-gdap-runner.git@v0.1.0",
# ]
# ///
from m365_gdap_runner import TenantContext

CLIENTS = ["graph", "exchange"]
SCOPES = ["User.Read.All"]
CONTINUE_ON_ERROR = True


def run(ctx: TenantContext) -> dict[str, object]:
    users = ctx.graph.get_all("/users", params={"$select": "userPrincipalName"})
    (config,) = ctx.exchange.run("Get-OrganizationConfig")
    return {"users": len(users), "audit_disabled": config["AuditDisabled"]}


if __name__ == "__main__":
    from m365_gdap_runner import main

    raise SystemExit(main(__name__))
```

The header at the top tells uv what the script needs: the runner, at a release, and anything
else the script imports, such as a library for its reports. `uv run script.py` builds the
script's own environment from it.

### What a script declares

- `run(ctx)`: the work for one tenant. `ctx` is the tenant's
  [context](#the-tenant-context); what `run()` returns is kept, as it is, in the tenant's
  [result](#reports).
- `CLIENTS`: which of `ctx.graph`, `ctx.exchange`, `ctx.ipps` and `ctx.arm` the script uses.
  Default `["graph"]`.
- `SCOPES`: the delegated Graph scopes the script needs in each tenant. Required when
  `graph` is in `CLIENTS`, not allowed otherwise. See [Scopes](#scopes).
- `report(results)`, optional: called once after the last tenant. See [Reports](#reports).
- `CONTINUE_ON_ERROR`, optional: keep going after an error in `run()`. See
  [Errors in run()](#errors-in-run).
- `GRANT_WHEN_POSSIBLE`, optional: add missing scopes to the app's grant where it can. See
  [Granting missing scopes](#granting-missing-scopes).

The script is checked before any sign-in, so a mistake in it never costs a browser prompt.

### The tenant context

`run(ctx)` gets a `TenantContext` for the tenant:

- `ctx.tenant`: the tenant, with its `tenant_id`, `display_name` and `cloud`.
- `ctx.graph`: Microsoft Graph.
- `ctx.exchange`: Exchange Online.
- `ctx.ipps`: Security & Compliance PowerShell.
- `ctx.arm`: Azure Resource Manager.
- `ctx.auth`: the azure-auth context for the tenant. Hand it to any azure-auth client, or to
  an Azure SDK client as `credential=`, for resources the four clients do not cover.

The clients come from [azure-auth](https://github.com/DrollRobot/azure-auth), which
documents what each one offers. Each client in `CLIENTS` is checked in the tenant before
`run()` starts, so `run()` never opens a sign-in window. Asking for a client the script did
not declare raises `ClientNotDeclared`.

### Scopes

The check before `run()` asks for exactly the scopes in `SCOPES`, so a tenant that has not
granted one of them fails before `run()`. `SCOPES = []` asks for none in particular: Graph
gets whatever the app holds in the tenant, and only a tenant where it holds nothing fails.
That suits a script that reads what the app holds, with `ctx.graph.granted_scopes()`.
`.default` is refused, since it would skip the check.

An entry may be a tuple of scopes, any one of which will do:

```python
SCOPES = [("Application.Read.All", "Application.ReadWrite.All")]
```

The check reads what each tenant has granted and asks for the first scope in the tuple
the tenant holds. A tenant that holds none is asked for the first scope in the tuple, so
the consent screen, the consent error and a self-grant name that one. Put the
least-privileged scope first. A tenant that signs in on its own is asked for the first
scope in each tuple, since what it has granted cannot be read before the sign-in.

### Flags of your own

`main(__name__)` reads the library's flags only. A script with flags of its own builds the
parser, adds the library's with `add_arguments()`, and passes what it parsed to `main()`;
`-h` then lists both:

```python
if __name__ == "__main__":
    import argparse
    from m365_gdap_runner import add_arguments, main

    parser = argparse.ArgumentParser(description="Users in each tenant.")
    parser.add_argument("--domain", help="Only users in this domain.")
    add_arguments(parser)
    args = parser.parse_args()
    DOMAIN = args.domain
    raise SystemExit(main(__name__, args))
```

`run()` reads `DOMAIN` like any other module-level value. The script's own flag names must
not clash with the library's.

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

## The tenants file

The tenants file is a CSV you edit in a spreadsheet, one row per customer. This package owns
seven columns, `m365_tenant_id`, `m365_tenant_name`, `m365_domain`, `m365_cloud`,
`m365_username`, `m365_gdap` and `m365_pw_url`, and reads one more per group a run names:
a column holding TRUE or FALSE, where a blank cell counts as FALSE and anything else is an
error. Every other column is left alone, so the same file can carry what other tools know
about the same customers on other platforms.

```csv
m365_tenant_id,m365_tenant_name,m365_domain,m365_cloud,m365_username,m365_gdap,m365_pw_url,pilot,patching,crm_id
aaaaaaaa-0000-0000-0000-00000000000a,Alpha,alpha.example,Commercial,,TRUE,,TRUE,FALSE,1001
bbbbbbbb-0000-0000-0000-00000000000b,beta corp,beta.example,USGov,admin@beta.us,,https://passwords.example/beta,TRUE,TRUE,1002
cccccccc-0000-0000-0000-00000000000c,Gamma,gamma.example,Commercial,admin@gamma.example,FALSE,,FALSE,TRUE,1003
```

Tenants are matched by ID; the name and domain columns are there for you. A row with no
`m365_tenant_id` is a customer without an M365 tenant and is skipped. A tenant may have more
than one row, as long as they agree in every `m365_` column; case and surrounding spaces do
not count, and other columns may differ. A run that selects a tenant whose rows disagree is
refused, naming the rows.

`m365_cloud` is `Commercial`, `USGov`, `USGovDoD` or `China`. `m365_username`, `m365_gdap`
and `m365_pw_url` are covered in [Tenants outside GDAP](#tenants-outside-gdap).

The file is `tenants.csv` in the config directory (`m365-gdap-runner-config path` shows
where), unless the `tenants_file` setting or `--tenants-file` points elsewhere.

### Excel workbooks

The file can also be an Excel workbook, named `.xlsx`, with the same columns in one table
(Insert > Table). The table may sit on any sheet, at any spot; a workbook with no table, or
more than one, is an error. TRUE and FALSE may be Excel's own or typed as text, a totals row
is ignored, and a formula counts as the value Excel last calculated.

When `m365-gdap-runner-customer-export` adds to the table, new columns go to its right and
new rows below it, and the table grows to hold them; every other cell is left as it was.
Those cells must be empty, and the table must have no totals row. Saving the workbook drops
any charts and images in it. A missing `.xlsx` file is created as a workbook with one table.

### Keeping it up to date

Only the `m365-gdap-runner-customer-export` command writes the file. It creates the file
when there is none. Any of the seven `m365_` columns the file lacks is added after the
`m365_` columns it has (or at the end when it has none), blank, except that `m365_gdap` is
TRUE for every tenant GDAP lists. A GDAP tenant with no row is appended with its name, its
cloud, TRUE in `m365_gdap` and every other column blank, and a warning names it. A blank
`m365_cloud` cell is filled in by OpenID Connect discovery. The command lists the tenants and
runs nothing, so run it once to create the file, then add group columns. Rows for tenants
without a GDAP relationship are yours to add and keep.

A script run never writes the file. Where the file falls behind, the run warns, names
`m365-gdap-runner-customer-export` as the fix, and goes on as if the command had run, for
that run only. With no file at all, it can pick only GDAP tenants.

## Sign-ins

### The partner sign-in

The runner signs in at the partner tenant as `username`, once per client in `CLIENTS`, since
each client uses a different Microsoft app; `client_id` names one app for all of them
instead.

Every selected tenant in the partner's cloud is then checked silently through that sign-in,
`jobs` at a time; every 30 seconds, a log line says how many are done and about how long the
rest will take. Each tenant's `run()` starts as soon as its access is settled.

Every sign-in, the partner's and each tenant's own, asks Microsoft for a new token for each
client rather than trusting one saved by an earlier run. A saved token says nothing about
whether Microsoft will still issue one: a Conditional Access policy may have started refusing
the sign-in since. So that failure happens in the check, where a window may open, instead of
inside `run()`, where none may.

Renewing your partner sign-in keeps the time you first signed in, and did multi-factor
authentication. A customer whose Conditional Access policy allows less time refuses it
however often it is renewed (AADSTS70043, or AADSTS50078 for the multi-factor
authentication). The first such customer in a run opens a new partner sign-in that asks for
your password, and multi-factor authentication, again: one window per client in `CLIENTS`.
That customer, and every later one refused for the same reason, is then checked through the
partner sign-in again. Only a tenant still refused goes to its own sign-in. Ctrl+C during
that window, then Enter, skips it, and those customers go to their own sign-in.

### Sign-in windows

Sign-ins open the system's default browser. The Windows sign-in window (the broker) is not
offered: GDAP does not work through it. To use another browser, or another profile of it:

- `browser_path`: the browser, as a full path (recommended) or a program on `PATH`.
- `browser_args`: a list of arguments for it, such as `["--profile-directory=Profile 2"]`.
  The sign-in address is added last. Needs `browser_path`.

A sign-in window waits `browser_timeout` seconds (default 120) to be completed. After that,
the sign-in fails: the partner sign-in ends the run, and a tenant's own leaves that tenant
`auth_failed`.

To give up on a tenant's own sign-in without waiting, press Ctrl+C, then Enter: that tenant
is `auth_failed` and the run goes on. Ctrl+C twice stops the run.

### Tenants outside GDAP

A tenant the partner sign-in cannot reach still runs. That covers a row of the tenants file
with no GDAP relationship, a tenant in another cloud (the partner sign-in reaches its own
cloud only), and a GDAP tenant where consent is missing or a policy wants a fresh sign-in.

Each tenant that fails the partner check, and each tenant in another cloud, gets a sign-in of
its own, in its own cloud, as the account in its `m365_username` cell: one sign-in per client
in `CLIENTS`. Sign-in windows open one at a time, as soon as a tenant needs one, while the
other checks go on. A tenant's `run()` never opens a sign-in window. With `--cache disk`,
later runs are silent. A sign-in that fails, is cancelled, times out or is skipped with
Ctrl+C then Enter leaves that tenant `auth_failed`; the rest run.

A tenant whose `m365_username` cell is blank gets no sign-in of its own: it is `auth_failed`,
with the reason, and no sign-in window opens for it.

`m365_pw_url` is optional: the address of the web page that holds that account's password,
such as its entry in a password manager. When the tenant's own sign-in opens a window, the
page opens right before it, once per tenant. It does not open when the sign-in comes from the
cache. It opens in the system's default browser, or in `pw_browser_path`, started with
`pw_browser_args`, such as `["--profile-directory=Default"]`; the address is added last.
Only `http` and `https` addresses are accepted.

The `m365_gdap` column says which tenants the partner sign-in reaches. TRUE, which every
appended row gets, tries the partner sign-in first. FALSE or blank never tries it: the tenant
goes straight to a sign-in of its own. GDAP can list a customer whose tenant it gives you no
access to; put FALSE in that row's cell. With `m365_username` blank as well, the tenant is
left out of the run, with a warning, instead of failing: that marks a customer you cannot
reach at all but cannot remove from Partner Center. Anything else in the cell is an error.

`m365_gdap` is a group like any other, so `--include m365_gdap` runs only the tenants marked
TRUE. A tenant there whose partner sign-in fails still gets a sign-in of its own when its
`m365_username` is filled in.

### Consent

A GDAP tenant whose grant lacks some of the script's scopes, where the app cannot add them
itself, goes to its own sign-in with Entra's consent screen, whatever the cache holds. Tick
"Consent on behalf of your organization" there. An administrator who consents for themselves
alone leaves the partner sign-in without the scopes, and Entra never shows the screen again
by itself for consent that exists. Granted for the organization, the tenant passes the
partner check on the next run.

### Granting missing scopes

A tenant whose grant lacks some of the script's `SCOPES` fails the check before the run. A
script that sets `GRANT_WHEN_POSSIBLE = True` has the check add them instead, where the app
holds `DelegatedPermissionGrant.ReadWrite.All` in that tenant. They go into the app's
tenant-wide grant, so every user of the app there gets them, each still limited by their own
roles. azure-auth logs each grant, which shows with `-vv` and in the debug log. The check
then waits for a token that carries the new scopes: about a minute, three at most. Where
the app does not hold that permission, the tenant gets a sign-in of its own as before.
`--grant-when-possible` and `--no-grant-when-possible` override the script.

### Keeping sign-ins between runs

Sign-ins are kept in memory and end with the run. To keep them, opt in to the disk cache:
the `cache` setting, or `--cache`, set to `disk` keeps them in an encrypted file, and later
runs skip the sign-in. That includes each tenant's own sign-in: a tenant the partner sign-in
cannot reach uses its cached sign-in, and opens the browser only when there is none.

The file is written after each sign-in, after each tenant, and when the run ends, so a run
that is killed keeps what was written before. A run that finds the file empty, which a run
killed while writing it could once leave, says so and names the file: delete it and sign in
again.

The file holds refresh tokens, which sign in without a password. Two settings help keep
track of it:

- `cache_path`: the file to keep it in. Empty uses azure-auth's file in the per-user cache
  directory.
- `cache_reminder_path`: a folder in plain sight, such as the Desktop. Every run with the
  disk cache puts `M365-GDAP-RUNNER_MSAL_CACHE_DELETEME` there, a link to the folder that
  holds the cache: a `.url` shortcut on Windows, a symbolic link elsewhere. It holds only a
  path, so it is safe in a synced folder; the cache itself is not.

## Running

### Tenants at once

A run works on one tenant at a time. The `jobs` setting changes that, and `--jobs` changes
it for one run; the silent checks before each `run()` use the same number. 5-8 suits
read-only reports; keep 1 for scripts that write unless they are idempotent.

### Errors in run()

An error raised in `run()` stops the run: that tenant is `error`, tenants already running
finish, and the tenants not yet started are `skipped`. That keeps a script that changes
things from repeating a broken change in every tenant. A script that only reads sets
`CONTINUE_ON_ERROR = True` to keep going. `--continue-on-error` and `--no-continue-on-error`
override the script.

A tenant that could not be reached is `auth_failed` and never stops the run.

### Stopping a run

Ctrl+C stops the run. A tenant it interrupts is `error`; with `jobs` above 1, the tenants
already running finish first. The tenants not yet started are `skipped`, and the summary is
printed. During a tenant's own sign-in window, Ctrl+C asks first: Enter skips that tenant,
and a second Ctrl+C stops the run.

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
stderr, so stdout is the script's.

### Logs

Each log line names the tenant it belongs to, or `-` outside one, so lines from tenants run
at once can be told apart. A script logs with Python's `logging` like any other program, and
its records go to the same place, named the same way. `-v` adds debug lines from the runner
and the script; `-vv` adds those of the libraries underneath, such as azure-auth and MSAL.

After the summary line, every warning and error the run logged is printed again, so none is
lost in a long run's output. That includes the script's own warnings, and each tenant that
failed, which logs an error.

A tenant's error, in the log and in its `TenantResult`, is one short line. For a tenant that
could not be reached, it names each sign-in tried: `partner: <why>; direct: <why>`, such as
`partner: missing 5 of 19 Graph scopes, self-grant failed; direct: no m365_username`. Missing
consent is counted, not listed. Microsoft's Trace ID, Correlation ID and Timestamp are left
out, and a message over 300 characters is cut to its first sentence.

The `debug_log_path` setting appends every log record to a file, debug and library ones
included, whatever `-v` says, and keeps every error in full. A folder there means
`m365_gdap_runner_debug.log` in it. Past `debug_log_max_mb` megabytes (default 10; 0 for no
limit), the file is compressed to `<name>.1.gz` and a new one started; the newest
`debug_log_keep` compressed files (default 5) are kept, and older ones deleted.

### Exit codes

- `0`: every tenant was `ok`.
- `1`: at least one tenant was not.
- `2`: the run could not start, or `report()` failed: a usage, configuration, script,
  sign-in, tenant listing, tenants file or tenant selection error.
- `130`: the run was stopped with Ctrl+C.

## Settings

Settings live in `config.toml` in the per-user config directory:

- Windows: `%LOCALAPPDATA%\m365_gdap_runner\config.toml`
- Linux: `~/.config/m365_gdap_runner/config.toml`
- macOS: `~/Library/Application Support/m365_gdap_runner/config.toml`

The `M365_GDAP_RUNNER_CONFIG_DIR` environment variable points at another directory instead.
The tenants file lives in the same directory unless `tenants_file` says otherwise.

### The config command

`m365-gdap-runner-config` creates and edits the file; it never needs editing by hand.

| Command | What it does |
|---|---|
| `init` | Asks for each setting in turn and writes the file. `--force` starts over. |
| `show` | Shows every setting's value and where it came from. |
| `set KEY VALUE` | Sets one setting. A list takes commas between its items. |
| `unset KEY` | Removes a setting, so its default applies. |
| `path` | Prints the file's path. |
| `open` | Opens the file in the program the system has for it. |
| `list-profiles` | Lists the profiles, marking the default. |
| `use PROFILE` | Makes a profile the default. |

### Every setting

| Setting | Default | What it is |
|---|---|---|
| `partner_tenant_id` | required | The partner tenant's ID or verified domain name. |
| `username` | required | The partner account that signs in. |
| `client_id` | empty | One app to sign in with for every client; empty uses Microsoft's first-party apps. |
| `cloud` | empty | The partner tenant's cloud: `Commercial`, `USGov`, `USGovDoD` or `China`; empty looks it up. |
| `tenants_file` | empty | The [tenants file](#the-tenants-file); empty uses `tenants.csv` in the config directory. |
| `jobs` | `1` | [Tenants at once](#tenants-at-once). |
| `cache` | `memory` | `memory`, or `disk` to [keep sign-ins between runs](#keeping-sign-ins-between-runs). |
| `cache_path` | empty | The disk cache's file; empty uses azure-auth's. |
| `cache_reminder_path` | empty | A folder for the cache's DELETEME link; empty puts none. |
| `browser_path` | empty | The browser [sign-in windows](#sign-in-windows) open in; empty uses the system default. |
| `browser_args` | empty | Arguments for `browser_path`. |
| `browser_timeout` | `120` | Seconds a sign-in window waits to be completed. |
| `pw_browser_path` | empty | The browser a tenant's `m365_pw_url` opens in; empty uses the system default. |
| `pw_browser_args` | empty | Arguments for `pw_browser_path`. |
| `debug_log_path` | empty | The [debug log](#logs) file or folder; empty writes none. |
| `debug_log_max_mb` | `10` | Megabytes the debug log grows to before it is compressed; 0 for no limit. |
| `debug_log_keep` | `5` | Compressed debug logs kept besides the current one. |

### Overrides

A flag beats an environment variable, which beats `config.toml`, which beats the default.
Every setting has an environment variable, its name upper-cased after
`M365_GDAP_RUNNER_`: `M365_GDAP_RUNNER_JOBS=5`. A list takes commas there too. These
settings also have a flag: `--partner-tenant-id`, `--username`, `--client-id`, `--cloud`,
`--tenants-file`, `--cache` and `--jobs`.

### Profiles

A profile is a named set of settings in the same file, for more than one partner tenant or
account. `m365-gdap-runner-config init --profile NAME` creates one, and `--profile NAME` on
`set`, `unset` and `show` works on it. A run uses `--profile`, else the
`M365_GDAP_RUNNER_PROFILE` environment variable, else the default profile that `use` sets.
