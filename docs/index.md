# m365-gdap-runner

Run a Python script across every customer tenant you reach through GDAP, from a single
partner sign-in. The runner owns the loop: sign-in, tenant selection, consent checks, error
isolation and parallelism. Authentication and the service clients come from
[azure-auth](https://github.com/DrollRobot/azure-auth). A script contains neither; it holds
the work for one tenant and its own reports.

- [Getting started](getting-started.md): install the commands, set up, and run a first
  script.
- [Usage](usage.md): writing scripts, choosing tenants, sign-ins, reports and settings.
- [Development](development.md): set up a development environment for the package itself.
- [Reference](reference/m365_gdap_runner.md): the API, from the docstrings.
