Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Ensure PS2EXE module is available for conversion
if (-not (Get-Command Invoke-PS2EXE -ErrorAction SilentlyContinue)) {
    Install-Module PS2EXE -Scope CurrentUser -Force
}
Import-Module PS2EXE

function Setup-RequiredModules {
    param ([string]$DistFolder)
    $modulesPath = Join-Path $DistFolder "Modules"
    $importExcelPath = Join-Path $modulesPath "ImportExcel"
    if (-not (Test-Path $modulesPath)) { New-Item -ItemType Directory -Path $modulesPath | Out-Null }
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) { Install-Module ImportExcel -Scope CurrentUser -Force }
    $installedModule = Get-Module -ListAvailable -Name ImportExcel | Select-Object -First 1
    $installedPath = $installedModule.ModuleBase
    if (Test-Path $importExcelPath) { Remove-Item $importExcelPath -Recurse -Force }
    Copy-Item -Path $installedPath -Destination $importExcelPath -Recurse
}

function Sign-Executable {
    param (
        [string]$ExePath,
        [string]$CertSha1
    )
    $signtool = Get-ChildItem -Path "C:\Program Files (x86)\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -match "\\x64\\signtool\.exe$" } |
                Sort-Object FullName -Descending | Select-Object -First 1
    if (-not $signtool) {
        [System.Windows.MessageBox]::Show("❌ signtool.exe not found.`nPlease install the Windows SDK to enable EXE signing.")
        return
    }
    try {
        & $signtool.FullName sign /fd SHA256 /sha1 $CertSha1 /tr http://timestamp.digicert.com /td SHA256 "$ExePath"
    } catch {
        [System.Windows.MessageBox]::Show("❌ Signing failed: $($_.Exception.Message)")
    }
}

# Load XAML UI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="PowerShell to EXE Converter" Height="560" Width="600">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="PowerShell Script (.ps1):" Margin="0,0,0,5" />
    <DockPanel Grid.Row="1">
      <TextBox x:Name="ScriptPathBox" Width="400" Margin="0,0,5,0" />
      <Button x:Name="BrowseScript" Content="Browse" Width="100" />
    </DockPanel>
    <TextBlock Grid.Row="2" Text="Program Name / Title Bar:" Margin="0,10,0,5" />
    <TextBox x:Name="ProgramNameBox" Grid.Row="3" Width="505" />
    <TextBlock Grid.Row="4" Text="Icon File (.ico) – Optional:" Margin="0,10,0,5" />
    <DockPanel Grid.Row="5">
      <TextBox x:Name="IconPathBox" Width="400" Margin="0,0,5,0" />
      <Button x:Name="BrowseIcon" Content="Browse" Width="100" />
    </DockPanel>
    <CheckBox x:Name="IncludeImportExcel" Grid.Row="6" Content="Bundle ImportExcel module (for Excel export)" Margin="0,15,0,0" />
    <GroupBox Grid.Row="7" Header="Security" Margin="0,15,0,10">
      <StackPanel>
        <Button x:Name="CheckSDK" Content="Check for Windows SDK (signtool.exe)" Width="250" Margin="0,5,0,5" />
        <CheckBox x:Name="SignExecutable" Content="Sign the EXE using a selected certificate" IsEnabled="False" Margin="0,5,0,5" />
        <ComboBox x:Name="CertSelector" DisplayMemberPath="Subject" Width="500" Margin="0,5,0,5" IsEnabled="False" />
        <DockPanel Margin="0,5,0,0">
          <TextBox x:Name="NewCertCN" Width="220" Margin="0,0,5,0" ToolTip="Common Name (CN)" />
          <ComboBox x:Name="CertExpiry" Width="80" Margin="0,0,5,0">
            <ComboBoxItem>30</ComboBoxItem><ComboBoxItem>90</ComboBoxItem>
            <ComboBoxItem>180</ComboBoxItem><ComboBoxItem IsSelected="True">365</ComboBoxItem>
            <ComboBoxItem>730</ComboBoxItem>
          </ComboBox>
          <Button x:Name="CreateCert" Content="Create Self-Signed Cert" Width="180" />
        </DockPanel>
        <DockPanel>
          <Button x:Name="ImportCert" Content="Import Certificate (PFX)" Width="200" Margin="0,5,0,0" />
          <Button x:Name="ExportCert" Content="Export Selected Certificate" Width="200" Margin="10,5,0,0" />
          <Button x:Name="DeleteCert" Content="Delete Self-Signed Cert" Width="200" Margin="10,5,0,0" />
        </DockPanel>
      </StackPanel>
    </GroupBox>
    <Button x:Name="ConvertScript" Grid.Row="8" Content="Convert to EXE" Height="30" Width="180" HorizontalAlignment="Center" Margin="0,20,0,0" />
  </Grid>
</Window>
"@

# Parse XAML
$reader = New-Object System.Xml.XmlTextReader([System.IO.StringReader]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Handlers
$window.FindName("BrowseScript").Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter = "PowerShell Scripts (*.ps1)|*.ps1";
    if ($dlg.ShowDialog()) { $window.FindName("ScriptPathBox").Text = $dlg.FileName }
})
$window.FindName("BrowseIcon").Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter = "Icon Files (*.ico)|*.ico";
    if ($dlg.ShowDialog()) { $window.FindName("IconPathBox").Text = $dlg.FileName }
})
$window.FindName("CheckSDK").Add_Click({
    $sign = $window.FindName("SignExecutable"); $certBox = $window.FindName("CertSelector");
    $sign.IsChecked = $false; $sign.IsEnabled = $false; $certBox.IsEnabled = $false
    $tool = Get-ChildItem "C:\Program Files (x86)\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
               Where-Object { $_.FullName -match "\\x64\\signtool\.exe$" } | Sort-Object FullName -Descending | Select-Object -First 1
    if ($tool) {
        $sign.IsEnabled = $true; $certBox.IsEnabled = $true; $certBox.Items.Clear();
        Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.EnhancedKeyUsageList.FriendlyName -contains 'Code Signing' } |
            ForEach-Object { $certBox.Items.Add([PSCustomObject]@{ Subject=$_.Subject; Thumbprint=$_.Thumbprint }) }
        [System.Windows.MessageBox]::Show("✅ SDK found. Certificates loaded.")
    } else {
        Start-Process "https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk/";
        [System.Windows.MessageBox]::Show("❌ signtool.exe not found. Please install SDK and re-check.")
    }
})
$window.FindName("CreateCert").Add_Click({
    $cn=$window.FindName("NewCertCN").Text; $days=[int]$window.FindName("CertExpiry").SelectedItem.Content
    if ([string]::IsNullOrWhiteSpace($cn)) { [System.Windows.MessageBox]::Show("Please enter a CN"); return }
    $cert=New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=$cn" -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddDays($days)
    [System.Windows.MessageBox]::Show("✅ Created: CN=$cn | Valid Until $($cert.NotAfter)")
    # Refresh and select new cert
    $cb=$window.FindName("CertSelector"); $cb.Items.Clear()
    Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.EnhancedKeyUsageList.FriendlyName -contains 'Code Signing'} |
        ForEach-Object { $cb.Items.Add([PSCustomObject]@{ Subject=$_.Subject; Thumbprint=$_.Thumbprint }) }
    $newItem=$cb.Items | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    $cb.SelectedItem = $newItem; $window.FindName("SignExecutable").IsChecked = $true
})
# Import certificate handler
$window.FindName("ImportCert").Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = "PFX Files (*.pfx)|*.pfx"
    if ($dlg.ShowDialog()) {
        $file = $dlg.FileName
        $pwd = Read-Host -AsSecureString "Enter password to protect/import PFX"
        try {
            $imported = Import-PfxCertificate -FilePath $file -CertStoreLocation Cert:\CurrentUser\My -Password $pwd -Exportable
            [System.Windows.MessageBox]::Show("✅ Imported: $($imported.Subject)")
            # Refresh list and select imported cert
            $cb = $window.FindName("CertSelector"); $cb.Items.Clear()
            Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.EnhancedKeyUsageList.FriendlyName -contains 'Code Signing'} |
                ForEach-Object { $cb.Items.Add([PSCustomObject]@{ Subject=$_.Subject; Thumbprint=$_.Thumbprint }) }
            $selItem = $cb.Items | Where-Object { $_.Thumbprint -eq $imported.Thumbprint }
            $cb.SelectedItem = $selItem; $window.FindName("SignExecutable").IsChecked = $true
        } catch {
            [System.Windows.MessageBox]::Show("❌ Import failed: $($_.Exception.Message)")
        }
    }
})
$window.FindName("ExportCert").Add_Click({
    $cb = $window.FindName("CertSelector"); $sel = $cb.SelectedItem
    if (-not $sel) { [System.Windows.MessageBox]::Show("Please select a certificate to export."); return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = "PFX Files (*.pfx)|*.pfx"; $dlg.FileName = ($sel.Subject -replace '[^\w]', '_') + ".pfx"
    if ($dlg.ShowDialog()) {
        $pwd = Read-Host -AsSecureString "Enter password to protect PFX file"
        try {
            Export-PfxCertificate -Cert Cert:\CurrentUser\My\$($sel.Thumbprint) -FilePath $dlg.FileName -Password $pwd -Force
            [System.Windows.MessageBox]::Show("✅ Certificate exported to: $($dlg.FileName)")
        } catch {
            [System.Windows.MessageBox]::Show("❌ Export failed: $($_.Exception.Message)")
        }
    }
})
$window.FindName("DeleteCert").Add_Click({
    $cb=$window.FindName("CertSelector"); $sel=$cb.SelectedItem
    if (-not $sel) { [System.Windows.MessageBox]::Show("Select a cert to delete."); return }
    if ([System.Windows.MessageBox]::Show("Delete selected cert?","Confirm",[System.Windows.MessageBoxButton]::YesNo) -eq [System.Windows.MessageBoxResult]::Yes) {
        Remove-Item "Cert:\CurrentUser\My\$($sel.Thumbprint)" -Force
        $cb.Items.Clear(); $window.FindName("SignExecutable").IsChecked=$false
        Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.EnhancedKeyUsageList.FriendlyName -contains 'Code Signing'} |
            ForEach-Object { $cb.Items.Add([PSCustomObject]@{ Subject=$_.Subject; Thumbprint=$_.Thumbprint }) }
        [System.Windows.MessageBox]::Show("Certificate deleted.")
    }
})
$window.FindName("ConvertScript").Add_Click({
    $script=$window.FindName("ScriptPathBox").Text; if(-not [System.IO.File]::Exists($script)){[System.Windows.MessageBox]::Show("Select valid script.");return}
    $name=$window.FindName("ProgramNameBox").Text; if([string]::IsNullOrWhiteSpace($name)){[System.Windows.MessageBox]::Show("Enter program name.");return}
    $outDir=Join-Path (Split-Path $script) "$name-Dist"; if(-not(Test-Path $outDir)){New-Item -ItemType Directory -Path $outDir|Out-Null}
    if($window.FindName("IncludeImportExcel").IsChecked){Setup-RequiredModules -DistFolder $outDir}
    $exe=Join-Path $outDir "$name.exe"
    try {
        Invoke-PS2EXE -InputFile $script -OutputFile $exe -IconFile $window.FindName("IconPathBox").Text -NoConsole
        if($window.FindName("SignExecutable").IsChecked){
            $sel=$window.FindName("CertSelector").SelectedItem; if(-not $sel){[System.Windows.MessageBox]::Show("Select cert.");return}
            Sign-Executable -ExePath $exe -CertSha1 $sel.Thumbprint
        }
        [System.Windows.MessageBox]::Show("✅ Conversion complete:`n$exe")
    } catch {
        [System.Windows.MessageBox]::Show("❌ Error:`n$($_.Exception.Message)")
    }
})

# Show window
$window.ShowDialog() | Out-Null
