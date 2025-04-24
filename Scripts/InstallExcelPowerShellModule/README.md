# PowerShell Script: Check and Install ImportExcel Module

## 📌 What This Script Does

This PowerShell script checks whether the `ImportExcel` module is already installed on your system.

- ✅ If it **is installed**, you'll get a confirmation message.
- ❌ If it **is not installed**, the script asks if you'd like to install it automatically.
- 💡 The `ImportExcel` module allows PowerShell to create and work with `.xlsx` Excel files (with multiple tabs, formatting, charts, etc.).

## 🧠 Why This Matters

PowerShell can export `.csv` files without any setup. But for `.xlsx` files, you need to install this module first — this script helps you do that easily.

---

## ▶️ How to Run This Script

1. **Open PowerShell ISE**
   - Press `Start`, type `PowerShell ISE`, and hit `Enter`.

2. **Paste the Script**
   - Copy the contents of the `Check-ImportExcel.ps1` file into the upper (script) pane.

3. **Press `F5` to Run**
   - This will execute the script. Follow the prompts.

---

## 🔐 Admin Rights Not Required

This installs the module **only for your user account**, so you **do not** need to run PowerShell as Administrator.

---

## 📥 Want to Install Manually?

If you prefer to install manually, just run this:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
