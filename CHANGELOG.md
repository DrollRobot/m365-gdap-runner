# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- A library for scripts that run across GDAP tenants. A script calls `main(__name__)` from
  its `__main__` block, and `uv run script.py` then signs in once per client it uses, lists
  GDAP tenants, and runs its `run(ctx)` per tenant with consent checks, error isolation and
  `--jobs` parallelism. `add_arguments()` adds the library's flags to a script's own parser,
  for scripts with flags of their own.
- Script contract: `run(ctx)`, the Graph `SCOPES` it needs (`[]` for none in particular,
  which gets whatever the app holds in each tenant), `CLIENTS` naming which of
  `ctx.graph`, `ctx.exchange`, `ctx.ipps` and `ctx.arm` it uses (default `graph`), and an
  optional `report(results)`, called once after the last tenant with every tenant's
  `TenantResult`, failed ones included. The library writes no output of its own: `run()`
  writes any per-tenant report and `report()` the run's.
- `TenantContext`: the tenant, its azure-auth context (`ctx.auth`, usable with any
  azure-auth or Azure SDK client), and one ready-made client per declared kind, each
  checked in the customer tenant before the first `run()` starts.
- Settings `partner_tenant_id`, `username`, `client_id`, `cache` (`memory`, or `disk` for
  an encrypted file) and `cloud`, with matching `--partner-tenant-id`, `--username`,
  `--client-id`, `--cache` and `--cloud` flags. Sign-ins open the browser; the Windows broker
  is not offered, since GDAP does not work through it.
- Settings `cache_path`, the file the disk cache is kept in (default: azure-auth's), and
  `cache_reminder_path`, a folder in plain sight where every run with the disk cache puts
  `M365-GDAP-RUNNER_MSAL_CACHE_DELETEME`, a link to the cache's folder, so the file is not
  forgotten.
- Settings `browser_path`, the browser sign-ins open in place of the system default, and
  `browser_args`, the arguments it is started with, such as a profile.
- Setting `browser_timeout`: the seconds a sign-in window waits to be completed before that
  sign-in fails (default 120).
- Ctrl+C during a tenant's own sign-in, then Enter, skips that tenant without waiting for
  `browser_timeout`: it is `auth_failed` and the run goes on. Ctrl+C twice stops the run.
- Setting `jobs`: how many tenants a run works on at once (default 1). `--jobs` overrides it
  for one run.
- Setting `debug_log_path`, for development: every log record, debug and library ones
  included, is appended to that file, or to `m365_gdap_runner_debug.log` in it when it names
  a folder. The terminal shows what it did without it.
- Tenant selection from a tenants file: a CSV, or an `.xlsx` workbook holding one table, with
  `m365_tenant_id`, `m365_tenant_name`, `m365_domain`, `m365_cloud`, `m365_username`,
  `m365_gdap`, `m365_pw_url`, and one TRUE/FALSE column per group; other tools may keep their
  own columns in the same file. Any of the `m365_` columns the file lacks is added, and a
  blank cloud is filled in by OpenID Connect discovery. A workbook's table grows to the right
  and down; the rest of the workbook is left as it was, except that charts and images are
  dropped.
  `--include` runs a tenant (display name or ID), a group (the tenants marked TRUE in that
  column) or `all` (every GDAP tenant); `--exclude` takes tenants and groups back out, and
  always wins. Both are repeatable, and a run that includes nothing is refused. A name that
  matches nothing is an error, and so is one that could mean two things: a tenant and a
  group, or `all` and either. A tenant or group may pick rows without a GDAP relationship,
  in any cloud. Only the `m365-gdap-runner` command writes the file: it creates it or adds
  what it lacks, warns which GDAP tenants it appended, and lists the tenants. A script run
  never writes it; it warns of the columns, GDAP tenants and clouds the file lacks, fills
  them in for itself, and names the command as the fix. Setting `tenants_file` and flag
  `--tenants-file` locate the file (default `tenants.csv` in the config directory).
- Tenants the partner sign-in cannot reach: before its `run()`, each selected tenant is
  checked silently through the partner sign-in, and each one that fails, or lives in
  another cloud, gets a sign-in of its own, in its own cloud, as its `m365_username`, one
  client at a time. Sign-in windows open one at a time, as soon as a tenant needs one, and
  each tenant's `run()` starts as soon as its access is settled, without waiting for the
  other tenants' checks and sign-ins; `run()` never prompts. Where `m365_username` is
  blank, those tenants are left `auth_failed` instead. `m365_gdap` marks the tenants the
  partner sign-in reaches: TRUE, which appended rows get, tries it first. FALSE or blank, as
  for a tenant GDAP lists but cannot reach, skips the silent check and goes straight to its
  own sign-in, or, with `m365_username` blank too, leaves the tenant out of the run with a
  warning. It is also a group, so `--include m365_gdap` runs only the tenants it marks.
  While the silent checks run, a log line every 30 seconds says how many are done and about
  how long the rest will take.
- Column `m365_pw_url`, optional: the web page that holds a tenant's own account's password.
  When that tenant's own sign-in opens a window, the page opens right before it, once per
  tenant, and not at all when the sign-in comes from the cache. Settings `pw_browser_path` and
  `pw_browser_args` choose the browser it opens in (default: the system default).
- Script setting `CONTINUE_ON_ERROR`: an error in `run()` stops the run unless the script sets
  it True, as the example reports and the service principal example do.
  `--continue-on-error` and the new `--no-continue-on-error` override it.
- Script setting `GRANT_WHEN_POSSIBLE`: where a tenant's grant lacks some of the script's
  `SCOPES` and the app holds `DelegatedPermissionGrant.ReadWrite.All` there, the check
  before the run adds them to the app's tenant-wide grant, through azure-auth's
  `grant_when_possible`, instead of failing the tenant. `--grant-when-possible` and
  `--no-grant-when-possible` override it.
- Example scripts, each with the header `uv run` builds its environment from:
  `examples/licensed_users.py` (Graph; its `report()` prints every tenant's users as JSON),
  `examples/audit_disabled.py` (Exchange Online; no `report()`, its log is the report),
  `examples/mailbox_audit_gaps.py` (mailboxes whose audit actions lag Microsoft's defaults,
  read through Exchange's REST endpoint; an Excel workbook per tenant and a summary
  workbook, in the folder its own `--out` flag names), `examples/unified_audit_log.py`
  (whether the unified audit log is on in each tenant; one Excel workbook for the run), and
  `examples/service_principal_enabled.py` (Graph; whether users can sign in to the
  enterprise apps whose name or app ID matches the regular expressions given, as one Excel
  workbook with a column per app; `--disable` turns sign-in off where it is on, and
  `--export-json` also writes the matched service principals, keyed by tenant ID).
- Supports Python 3.11 and newer. CI tests every supported version, 3.11 through 3.14, on
  Ubuntu, macOS and Windows.

### Changed

- Authentication, token caching, tenant lookup, and the Graph, Exchange Online, Security &
  Compliance and Azure Resource Manager clients come from
  [azure-auth](https://github.com/DrollRobot/azure-auth), pinned to v0.5.0. This package
  no longer carries an MSAL session or a Graph client of its own.

[Unreleased]: https://github.com/DrollRobot/m365_gdap_runner/commits/HEAD
