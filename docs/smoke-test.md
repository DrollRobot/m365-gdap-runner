# Smoke test

Manual checks against a real partner tenant. The automated suite is offline
only; run these once signing in to a GDAP partner tenant is possible.

## Setup

```bash
uv sync
uv run m365-gdap-runner-config init      # sets partner_tenant_id and username
```

## 1. List tenants

```bash
uv run m365-gdap-runner --list-tenants
```

Expect one browser sign-in (Microsoft Graph Command Line Tools), then one
`tenant_id<TAB>display_name` line per GDAP customer.

## 2. Graph payload against one tenant

```bash
uv run m365-gdap-runner examples/licensed_users.py --tenant "<name or ID>"
```

Expect a JSON array with one record: `status` `ok` with licensed users in
`result` (exit code 0), or `auth_failed` naming `AADSTS65001` if
`User.Read.All` is not consented in that tenant (exit code 1).

## 3. Exchange payload against one tenant

```bash
uv run m365-gdap-runner examples/audit_disabled.py --tenant "<name or ID>"
```

Expect a second browser sign-in, for Exchange Online PowerShell, then one record
with `name` and `auditDisabled`. This is the first live run of azure-auth's
Exchange client through a GDAP sibling context; an `auth_failed` here means the
customer tenant refused the silent Exchange token.

## 4. Standalone

```bash
uv run python examples/licensed_users.py --tenant "<name or ID>"
```

Expect the same users as step 2, printed as JSON.

## 5. Disk token cache

```bash
uv run m365-gdap-runner --list-tenants --cache disk
uv run m365-gdap-runner --list-tenants --cache disk
```

Expect the first run to open the browser and the second not to. The cache is
azure-auth's `msal_token_cache.bin` in the per-user cache directory, encrypted,
not readable as plain text.

## 6. Parallel run to CSV

```bash
uv run m365-gdap-runner examples/licensed_users.py --jobs 5 --format csv --output users.csv
```

Expect stderr log lines tagged with each tenant's name, a final `Summary:`
line, and `users.csv` rows starting with `tenant_id` and `display_name`.
