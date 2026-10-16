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
- Script contract: `run(ctx)`, the Graph `SCOPES` it needs, `CLIENTS` naming which of
  `ctx.graph`, `ctx.exchange`, `ctx.ipps` and `ctx.azure` it uses (default `graph`), and an
  optional `report(results)`, called once after the last tenant with every tenant's
  `TenantResult`, failed ones included. The library writes no output of its own: `run()`
  writes any per-tenant report and `report()` the run's.
- `TenantContext`: the tenant, its azure-auth context (`ctx.auth`, usable with any
  azure-auth or Azure SDK client), and one ready-made client per declared kind, each
  checked in the customer tenant before the first `run()` starts.
- Settings `partner_tenant_id`, `username`, `client_id`, `cache` (`memory`, or `disk` for
  an encrypted file) and `cloud`, with matching `--partner-tenant-id`, `--username`,
  `--client-id`, `--cache` and `--cloud` flags.
- Tenant selection from a tenants file: a CSV with `m365_tenant_id`, `m365_tenant_name`,
  `m365_domain`, `m365_cloud`, an optional `m365_username`, and one TRUE/FALSE column per
  group; other tools may keep their own columns in the same file. A blank cloud is filled in
  by OpenID Connect discovery. `--group NAME` runs the tenants marked in that column
  (repeatable, combines with `--tenant`), `--all` runs every GDAP tenant, and a run that
  names neither `--group`, `--tenant` nor `--all` is refused. `--group` and `--tenant` may
  pick rows without a GDAP relationship, in any cloud. Every run appends the GDAP tenants the
  file lacks and warns which; the `m365-gdap-runner` command creates or refreshes the file and
  lists the tenants. Setting `tenants_file` and flag `--tenants-file` locate the file (default
  `tenants.csv` in the config directory).
- Tenants the partner sign-in cannot reach: before the first `run()`, every selected tenant
  is checked silently through the partner sign-in, and each one that fails, or lives in
  another cloud, gets a sign-in of its own, in its own cloud, as its `m365_username` (the
  partner account when blank), one client at a time, so the loop never prompts.
  `--gdap-only` leaves those tenants `auth_failed` instead.
- Example scripts, each with the header `uv run` builds its environment from:
  `examples/licensed_users.py` (Graph; its `report()` prints every tenant's users as JSON),
  `examples/audit_disabled.py` (Exchange Online; no `report()`, its log is the report), and
  `examples/mailbox_audit_gaps.py` (mailboxes whose audit actions lag Microsoft's defaults,
  read through Exchange's REST endpoint; an Excel workbook per tenant and a summary
  workbook, in the folder its own `--out` flag names).

### Changed

- Authentication, token caching, tenant lookup, and the Graph, Exchange Online, Security &
  Compliance and Azure Resource Manager clients come from
  [azure-auth](https://github.com/DrollRobot/azure-auth), pinned to v0.4.0. This package
  no longer carries an MSAL session or a Graph client of its own.

[Unreleased]: https://github.com/DrollRobot/m365_gdap_runner/commits/HEAD
