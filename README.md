# NetBackup Services Tool

Phase 1 desktop tool for discovering NetBackup clients from a selected Master Server and testing connectivity through that Master Server.

## Build in GitHub

The workflow file must be located here:

```text
.github/workflows/build-windows-exe.yml
```

Then go to **Actions** and run **Build Windows EXE**.

## Run

Launch `NetBackup_Services_Tool.exe` as the CyberArk/service account, for example `greif\p-ca-dls`.

The app does not store credentials. PowerShell commands run under the account used to launch the EXE.
