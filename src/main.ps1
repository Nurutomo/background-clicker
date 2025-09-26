$global:actions = @()

. "$PSScriptRoot\win32.ps1"
. "$PSScriptRoot\functions.ps1"
. "$PSScriptRoot\gui.ps1"

$Global:gui.Add_Load({ Update-WindowList })
$Global:gui.Add_Load({
  Set-RecentFiles(Get-RecentFiles)
})

# Setup Ctrl+C handling before showing GUI
$null = [Console]::TreatControlCAsInput = $false

# Register cleanup event
Register-EngineEvent PowerShell.Exiting -Action {
  Write-Host "Exiting AutoClicker-Background..."
  if ($gui -and !$gui.IsDisposed) {
    $gui.Close()
    $gui.Dispose()
  }
} | Out-Null

# Set up a flag for graceful shutdown
$Global:ShouldExit = $false

# Show GUI non-modally and keep console responsive
$gui.Show()

# Keep the script running while GUI is open
try {
  while ($gui.Visible -and !$Global:ShouldExit) {
    [System.Windows.Forms.Application]::DoEvents()
    
    # Check for Ctrl+C by monitoring console key availability
    if ([Console]::KeyAvailable) {
      $key = [Console]::ReadKey($true)
      if ($key.Key -eq [ConsoleKey]::C -and $key.Modifiers -band [ConsoleModifiers]::Control) {
        Write-Host "`nShutting down AutoClicker-Background..."
        $Global:ShouldExit = $true
        break
      }
    }
    
    Start-Sleep -Milliseconds 50
  }
}
catch {
  Write-Host "Application interrupted: $($_.Exception.Message)"
}
finally {
  if ($gui -and !$gui.IsDisposed) {
    $gui.Close()
    $gui.Dispose()
  }
}