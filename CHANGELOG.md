# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.2.0] - 2026-10-29 - azure-auth 0.10.0 network errors

### Changed

- **Breaking:** azure-auth v0.10.0. `ctx.graph`, `ctx.exchange`, `ctx.ipps`, `ctx.arm` and
  `ctx.auth` raise `azure_auth.NetworkError` when a host cannot be reached, in place of
  `httpx` and `requests` exceptions, and check TLS against the OS trust store. `main()`'s
  `transport=` takes an `httpx2` transport.

## [0.1.0] - 2026-10-29 - Initial release

### Added

- A library for running a Python script across GDAP tenants: it signs in as the partner, checks
  access to each tenant, and runs the script per tenant, one at a time or in parallel.
- Tenant selection by name, ID or group, from a CSV or Excel tenants file.
- Perfers signing in with GDAP permissions, but falls back to a direct sign in if no access or
  missing scopes.
- Missing Graph consent handled at the consent screen or, if
  `DelegatedPermissionGrant.ReadWrite.All` is granted, adding scopes automatically.
- An optional encrypted token cache on disk.
- Support for Python 3.11 and newer on Linux, macOS and Windows.

[Unreleased]: https://github.com/DrollRobot/m365-gdap-runner/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/DrollRobot/m365-gdap-runner/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/DrollRobot/m365-gdap-runner/releases/tag/v0.1.0
