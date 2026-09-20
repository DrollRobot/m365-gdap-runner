# Smoke test

Manual checks against a real partner tenant. The automated suite is offline
only; run these once signing in to the configured app is possible.

## Setup

```bash
uv sync
uv run m365-gdap-runner-config init      # sets partner_tenant_id
```

## 1. List tenants

```bash
uv run m365-gdap-runner --list-tenants
```

Expect one browser sign-in, then one `tenant_id<TAB>display_name` line per
customer with an active GDAP relationship.

## 2. Example payload against one tenant

```bash
uv run m365-gdap-runner examples/licensed_users.py --tenant "<name or ID>"
```

Expect a JSON array with one record: `status` `ok` with licensed users in
`result` (exit code 0), or `auth_failed` naming `consent_required` if
`User.Read.All` is not consented in that tenant (exit code 1).

## 3. Standalone

```bash
uv run python examples/licensed_users.py --tenant "<name or ID>"
```

Expect the same users as step 2, printed as JSON.

## 4. Persisted token cache

```bash
uv run m365-gdap-runner --list-tenants --persist-token-cache
uv run m365-gdap-runner --list-tenants --persist-token-cache
```

Expect the first run to open the browser and the second not to. The cache is
`msal_token_cache.bin` in the config directory and is not readable as plain text.

## 5. Parallel run to CSV

```bash
uv run m365-gdap-runner examples/licensed_users.py --jobs 5 --format csv --output users.csv
```

Expect stderr log lines tagged with each tenant's name, a final `Summary:`
line, and `users.csv` rows starting with `tenant_id` and `display_name`.
