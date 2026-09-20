# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `m365-gdap-runner` CLI: one partner sign-in, GDAP tenant enumeration, and a payload's
  `run(ctx)` executed per tenant with consent checks, error isolation, `--jobs`
  parallelism, and JSON/CSV output.
- Payload contract (`SCOPES` + `run(ctx)`), file and entry-point loading, and
  `standalone()` for running a payload against one tenant.
- `GraphClient` with per-call v1.0/beta selection, paging, and throttling retries.
- In-memory token cache by default; opt-in encrypted cache (`--persist-token-cache`).
- Example payload `examples/licensed_users.py` and `docs/smoke-test.md`.

[Unreleased]: https://github.com/DrollRobot/m365_gdap_runner/commits/HEAD
