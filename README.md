# Rexie Tools

PowerShell automation toolkit for endpoint support workflows.

---

## Features

- OS-aware credential storage (Windows + macOS)
- GitHub version checking
- Windows standalone EXE build
- macOS portable launcher

---

## Installation

### Windows

1. Go to the **Releases** page.
2. Download the latest `RexieTools.exe`.
3. Run the executable.

No installation required.

---

### macOS

1. Install PowerShell 7.
2. Download the repository.
3. Run:

./run_rexietools.command

---

## Versioning

Releases follow semantic versioning:

vMajor.Minor.Patch

The application checks GitHub for newer versions and will notify users when updates are available.

---

## Security

Credentials are stored locally using OS-native secure mechanisms:

- Windows: Secure credential storage
- macOS: Local encrypted credential storage

No credentials are stored in this repository.

---

## License

Internal use and distribution permitted.
