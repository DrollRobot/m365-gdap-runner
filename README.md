# m365-gdap-runner

[![CI](https://github.com/DrollRobot/m365-gdap-runner/actions/workflows/ci.yml/badge.svg)](https://github.com/DrollRobot/m365-gdap-runner/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.11%2B-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Run a Python script across every customer tenant you reach through GDAP, from a single
partner sign-in. The runner handles the sign-ins, tenant selection, consent checks, error
isolation and parallelism; a script holds only the work for one tenant and its own reports.

See the [documentation](docs/index.md).

## License

MIT; the full text is in [LICENSE](LICENSE).
