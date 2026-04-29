# NetBackup_Service_Tool

Phase 1 desktop tool for discovering NetBackup clients from a selected Master Server and testing connectivity through that Master Server.

## Build in GitHub

1. Create a GitHub repo.
2. Upload:
   - `app.py`
   - `requirements.txt`
   - `.github/workflows/build-windows-exe.yml`
3. Go to **Actions**.
4. Run **Build Windows EXE**.
5. Download the `NetBackup_Service_Tool` artifact.
6. Extract the ZIP to get `NetBackup_Service_Tool.exe`.

## Run

Launch the EXE as the CyberArk/service account, for example `greif\p-ca-dls`.

The app does not store credentials. PowerShell commands run under the account used to launch the EXE.
