# SMTP Diagnostic Tool - GitHub Documentation

Welcome to the **SMTP Diagnostic Tool** documentation! This tool is built using **PowerShell 7+** and is intended to provide detailed diagnostics and debug logging for SMTP connections — including Office365, AzureComm, and other STARTTLS/SSL-based SMTP services.

---

## 🔧 Prerequisites

- PowerShell **7.0 or higher** (cross-platform)
- Windows OS for GUI version (uses `System.Windows.Forms`)
- Optional: .NET SDK for debugging or modification

> ❗ **Note:** The graphical user interface (GUI) version **only works on Windows**. For Linux/macOS, use a CLI-only version.

---

## 📦 Features

- GUI built with WinForms for simple operation
- Raw TCP + SSL connection testing
- STARTTLS and Implicit SSL support
- TLS handshake details (protocol, cipher, cert chain)
- Full SMTP command transcript logging
- Built-in test email sending
- Toggleable debug categories:
  - SMTP Connection Info
  - SMTP Protocol Commands
  - TLS Handshake Details

---

## 🗂 Project Structure

```
.
├── smtp-diagnostic-tool.ps1      # Main PowerShell script
├── README.md                     # Main documentation
└── LICENSE                       # GNU General Public License v3.0
```

---

## ✅ Best Practices

### 🔐 Security

- Avoid committing credentials or secrets.
- Use [PowerShell Secrets Management](https://learn.microsoft.com/en-us/powershell/secrets/) to store credentials securely.
- When sharing logs, redact sensitive info (username, auth headers).

### 🧪 Testing

- Use PowerShell 7+ with `-NoProfile` for clean runs.
- Validate Office365/AzureComm compatibility using the full debug test email button.

### 📝 Logging

- All debug logs include timestamps.
- Raw SMTP commands/responses are labeled:
  - `CLIENT>>` – what the client sends
  - `SERVER>>` – what the SMTP server responds

### 💻 Cross-Platform Support

- GUI version: **Windows only** (WinForms)
- Core logic: TCP/SSL/SMTP works cross-platform
- CLI wrapper for Linux/macOS planned (see Roadmap)

### 📁 Versioning

Use SemVer-style versioning, tagged in Git:

```
v1.0.0 - Initial working version with full debug
v1.1.0 - Enhanced TLS chain logging, improved layout
v1.2.0 - CLI support for Linux/Mac (planned)
```

---

## 🚀 Getting Started

```bash
pwsh
./smtp-diagnostic-tool.ps1
```

> Run with PowerShell 7+. A GUI will appear (on Windows). It might be hidden behind the powershell icon in the taskbar.

---

## 🧭 Roadmap

If there's any desire for improvement, hit us up, and your query might end up on the comitted roadmap.
- Nothing on the roadmap right now.

---

## 🤝 Contributing

Pull requests welcome!

- Fork the repo
- Create a new branch
- Submit PR with meaningful title

---

## 📄 License

GNU General Public License v3.0

---

## 🙋 FAQ

**Q: Does it support Gmail SMTP?**

> Yes, but you must use an App Password or OAuth2 (for Gmail).

**Q: Does this work on Linux?**

> The backend logic does — the GUI does not.

**Q: Can I test arbitrary SMTP servers?**

> Yes — just change the hostname/port.

---

## 📬 Contact

For questions or bugs, please open an issue on GitHub.

