# 🔐 Word Document Self-Encryption with VBA

This project demonstrates how to build and share **self-encrypting Microsoft Word documents** using **VBA macros**. The system enables documents to:
- Be opened a limited number of times
- Automatically encrypt after a set period
- Restrict access by date, location, or user
- Log access attempts

> ⚠️ For Microsoft Word desktop (not Word Online). Macros must be enabled.

---

## 📁 Project Structure

| File | Description |
|------|-------------|
| `secure_document.docm` | The macro-enabled Word document |
| `self_encrypting.vba` | VBA source code for encryption logic |
| `readme.txt` | Instructions for recipients |
| `PrepareDocumentForSharing` macro | Resets usage counters before distribution |

---

## 🧠 How It Works

- **Open Count Limit**: Document can only be opened `X` times.
- **Auto-Encrypt Timer**: Document encrypts after `Y` seconds.
- **VBA Macro**: Embedded in the `.docm` file, handles logic.
- **Optional Conditions**: Expiry dates, computer name check, etc.

---

## 🚀 Quick Start

### 1. Setup

1. Open `secure_document.docm` in Microsoft Word.
2. Press `Alt + F11` to open the VBA editor.
3. Paste the contents of `self_encrypting.vba` into `ThisDocument`.
4. Save the document as `.docm`.

### 2. Modify Configuration

In the VBA script:

```vba
Const maxOpens = 3
Const autoEncryptDelay = 60 ' in seconds
Add optional constraints:


Dim expiryDate As Date
expiryDate = DateValue("2024-12-31")

If Date > expiryDate Then
    MsgBox "This document has expired."
    ActiveDocument.Close SaveChanges:=False
End If