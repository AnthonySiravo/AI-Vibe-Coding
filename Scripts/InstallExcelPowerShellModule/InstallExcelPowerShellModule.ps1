# Check if the ImportExcel module is available
if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "✅ ImportExcel module is already installed."
} else {
    Write-Host "❌ ImportExcel module is not installed."

    # Ask the user if they want to install it
    $install = Read-Host "Would you like to install it now? (Y/N)"
    if ($install -eq "Y") {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
        Write-Host "✅ ImportExcel module has been installed."
    } else {
        Write-Host "You can install it later using:"
        Write-Host "Install-Module -Name ImportExcel -Scope CurrentUser -Force"
    }
}
