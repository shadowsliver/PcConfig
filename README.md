# PcConfig

## ⚙️ Requirements

- This script must be run as **Administrator**.
- The PowerShell `ExecutionPolicy` should be set to one of the following:
  - `RemoteSigned`
  - `Bypass`
  - `Unrestricted`

---

## 🛠️ Troubleshooting

If there are issues with the script:

1. Open PowerShell as Administrator.
2. Navigate to the script folder using:

   ```powershell
   cd <path-to-folder>
   # or
   Set-Location <path-to-folder>
   ```

3. Run the script:

   ```powershell
   .\Configure-ProCom.ps1
   ```

---

## ⚡ Quick Mode

Quick mode applies a predefined set of configurations without user interaction.

---

## 🐞 Winget Issues

If Winget is not working properly:

- Use **debug mode** to fix it.
- Opening the **Windows Store page** and updating via the **debug menu (option 3)** usually resolves the issue.

---

## 📦 Extra Installations

Place additional installers in the `install` folder next to the script. Configure them using:

```plaintext
.\Install\installations.csv
```

**Example format:**

```
software;file;arguments
SOFTWARENAME;FILENAME.EXE;ARGUMENTS
SOFTWARENAME;FILENAME.EXE;ARGUMENTS
SOFTWARENAME;FILENAME.EXE;ARGUMENTS
```

- `arguments` can be left empty.

---

## 🧩 Extra Configurations

Additional configurations can be placed in the `config` folder.

### Network Drives

```plaintext
.\config\netdrive.csv
```

**Example:**

```
Drive;Path
N;\\server\data
B;\\test\Folder1
```

### User Accounts

```plaintext
.\config\users.csv
```

**Example:**

```
user;fullname;password
JDoe;John Doe;Password123
```