# PowerShell Script: Convert PowerShell Scripts to EXE (with Optional Module Bundling)

## 📌 What This Script Does

This PowerShell script launches a **Graphical User Interface (GUI)** that helps you:
- 📝 Select a `.ps1` script file
- 🏷️ Specify a program name (used for the EXE filename and title bar)
- 🎨 Optionally select an `.ico` file to brand your EXE
- 📦 Bundle the **ImportExcel** module so the EXE can generate `.xlsx` files without extra installs
- 🔍 Detect the Windows SDK (`signtool.exe`) and enable code-signing controls
- 🔐 Create self-signed code-signing certificates, auto-select them, and check the “Sign the EXE” option
- 📥 Import a paid or PFX certificate into your user store and auto-select it for signing
- 📤 Export the selected certificate as a PFX (protected by your chosen password)
- 🗑️ Delete unwanted self-signed certificates
- ⚙️ Compile your script into an EXE using the **PS2EXE** module
- 📝 Produce an output folder named `<ProgramName>-Dist` containing:
  - Your signed (or unsigned) `.exe`
  - A `Modules/ImportExcel` folder if bundling was selected

---

## 🧠 Why This Matters

- **User-friendly distribution**: Provide non-technical users with a double‑clickable EXE  
- **All-in-one package**: No need for end‑users to install prerequisites like ImportExcel or PS2EXE  
- **Trust and professionalism**: Add a custom icon and digitally sign your EXE so it’s recognized by Windows and security tools  

---

## ▶️ How to Use

1. **Launch the GUI**  
   ```powershell
   & "ConvertToExe.ps1"
   ```
2. **Script Selection**  
   Click **Browse** and choose your `.ps1` file.  
3. **Program Name**  
   Enter the name for your EXE; this also names the `<ProgramName>-Dist` output folder.  
4. **Icon (Optional)**  
   Browse for an `.ico` file to include.  
5. **Bundle ImportExcel**  
   Check the box if your script writes `.xlsx` files.  
6. **Check for Windows SDK**  
   Click **Check for Windows SDK**. If found, the signing controls will enable.  
7. **Manage Certificates**  
   - **Create Self‑Signed Cert**: Generates, selects it, and checks “Sign the EXE.”  
   - **Import Certificate (PFX)**: Imports a paid or external certificate into your CurrentUser\My store and auto-selects it.  
   - **Export Selected Certificate**: Opens a save dialog to export a PFX (you’ll set a password).  
   - **Delete Self‑Signed Cert**: Removes unwanted certs from `Cert:\CurrentUser\My`.  
8. **Sign the EXE**  
   Check “Sign the EXE” to embed your certificate during conversion.  
9. **Convert to EXE**  
   Click **Convert to EXE**. The tool will:
   - Create `<ProgramName>-Dist` next to your script  
   - Bundle ImportExcel if selected  
   - Generate the `.exe` via PS2EXE  
   - Sign it if requested  
   - Show a confirmation with the output path  

---

## 🔐 Requirements

- **PowerShell 5.1+** (with WPF support)  
- **Windows 10 SDK** (`signtool.exe`) for signing  
- **Internet access** for automatic module installation  

---

## ⚙️ Manual Setup

If you prefer manual installation:

```powershell
# PS2EXE module
Install-Module -Name PS2EXE -Scope CurrentUser -Force

# ImportExcel module
Install-Module -Name ImportExcel -Scope CurrentUser -Force

# Download Windows 10 SDK (contains signtool.exe):
# https://developer.microsoft.com/windows/downloads/windows-10-sdk/
```

---

## 📝 Troubleshooting

- **“signtool.exe not found”**: Install/repair the Windows 10 SDK and re-run **Check for Windows SDK**.  
- **“Select a certificate”**: Ensure you’ve created, imported, or paid for a Code Signing cert under `Cert:\CurrentUser\My`.  
- **Conversion errors**: Review the popup for details—common issues include missing PS2EXE or invalid paths.

> _This tool is provided as-is without warranty. Use at your own risk._
