# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `m365-gdap-runner` CLI: one partner sign-in per client the payload uses, GDAP tenant
  listing, and a payload's `run(ctx)` executed per tenant with consent checks, error
  isolation, `--jobs` parallelism, and JSON/CSV output.
- Payload contract: `run(ctx)`, the Graph `SCOPES` it needs, and `CLIENTS` naming which of
  `ctx.graph`, `ctx.exchange`, `ctx.ipps` and `ctx.azure` it uses (default `graph`). File
  and entry-point loading, and `standalone()` for running a payload against one tenant.
- `TenantContext`: the tenant, its azure-auth context (`ctx.auth`, usable with any
  azure-auth or Azure SDK client), and one ready-made client per declared kind, each
  checked in the customer tenant before `run()` starts.
- Settings `partner_tenant_id`, `username`, `client_id`, `cache` (`memory`, or `disk` for
  an encrypted file) and `cloud`, with matching `--partner-tenant-id`, `--username`,
  `--client-id`, `--cache` and `--cloud` flags.
- Tenant selection from a tenants file: a CSV with `m365_tenant_id`, `m365_tenant_name`,
  `m365_domain`, `m365_cloud` and one TRUE/FALSE column per group; other tools may keep
  their own columns in the same file. The cloud is filled in by OpenID Connect discovery,
  and a group member in a cloud other than the run's is skipped. `--group NAME` runs the
  tenants marked in that column
  (repeatable, combines with `--tenant`), `--all` runs every GDAP tenant, and a run that
  names neither `--group`, `--tenant` nor `--all` is refused. Every run appends the GDAP
  tenants the file lacks and warns which; `--list-tenants` creates or refreshes the file
  without running anything. Setting `tenants_file` and flag `--tenants-file` locate the file
  (default `tenants.csv` in the config directory).
- Example payloads `examples/licensed_users.py` (Graph) and `examples/audit_disabled.py`
  (Exchange Online).

### Changed

- Authentication, token caching, tenant lookup, and the Graph, Exchange Online, Security &
  Compliance and Azure Resource Manager clients come from
  [azure-auth](https://github.com/DrollRobot/azure-auth), pinned to v0.3.0. This package
  no longer carries an MSAL session or a Graph client of its own.

[Unreleased]: https://github.com/DrollRobot/m365_gdap_runner/commits/HEAD
