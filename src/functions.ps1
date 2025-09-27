# Functions file - contains utility functions
# Dependencies are loaded by main.ps1

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Initialize required script-level variables
$script:isRecording = $false
$script:recordedEvents = @()
$script:recordStartTime = 0
$script:keyboardHook = [IntPtr]::Zero
$script:mouseHook = [IntPtr]::Zero
$script:keyStates = @{}
$script:isPlaying = $false
$script:playbackTimer = New-Object System.Windows.Forms.Timer
$script:playbackIndex = 0
$script:playbackLoopCount = 1
$script:playbackCurrentLoop = 0
$script:pressedKeys = @{}
$script:pressedMouseButtons = @{}
$script:targetHwnd = [IntPtr]::Zero


##### File systems #####

# New
$script:currentFile = $null
function New-AutomationScript {
  $script:currentFile = $null
  $script:actions = @()
  $Global:statusLabel.Text = "New Automation Created"
  Update-EventList
}

# Open
function Open-AutomationScript {
  $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $openFileDialog.Filter = "JSON files (*.json)|*.json"
  $openFileDialog.Title = "Load Actions"
  if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    try {
      $jsonContent = Get-Content $openFileDialog.FileName -Raw
      $script:actions = ConvertFrom-Json $jsonContent
      Update-EventList
      Add-RecentFile $openFileDialog.FileName
      $script:currentFile = $openFileDialog.FileName
      $statusLabel.Text = "Actions loaded from $($openFileDialog.FileName)"
    }
    catch {
      [System.Windows.Forms.MessageBox]::Show("Error loading file: $($_.Exception.Message)", "Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
  }
}

# Save
function Save-AutomationScript {
  if ($null -eq $script:currentFile) {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "JSON files (*.json)|*.json"
    $saveFileDialog.Title = "Save Actions"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $jsonContent = $script:actions | ConvertTo-Json -Depth 3
      $script:currentFile = $saveFileDialog.FileName
      $statusLabel.Text = "Actions saved to $($saveFileDialog.FileName)"
    }
  }
  if ($null -eq $script:currentFile) {
    return
  }
  try {
    $jsonContent = $script:actions | ConvertTo-Json -Depth 5
    Set-Content -Path $script:currentFile -Value $jsonContent -Encoding UTF8
    Add-RecentFile $script:currentFile
    $statusLabel.Text = "Actions saved to $script:currentFile"
    return
  }
  catch {
    [System.Windows.Forms.MessageBox]::Show("Error saving file: $($_.Exception.Message)", "Save Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
  }
}

# Recent Files
function Get-RecentFiles {
  try {
    $registryPath = "HKCU:\Software\AutoClicker\RecentFiles"
    if (Test-Path $registryPath) {
      $recentFiles = @()
      for ($i = 1; $i -le 10; $i++) {
        $file = Get-ItemProperty -Path $registryPath -Name "File$i" -ErrorAction SilentlyContinue
        if ($file -and $file."File$i" -and (Test-Path $file."File$i")) {
          $recentFiles += $file."File$i"
        }
      }
      return $recentFiles
    }
  }
  catch {
    Write-Warning "Error reading recent files from registry: $($_.Exception.Message)"
  }
  return @()
}

function Add-RecentFile {
  param([string]$FilePath)
  
  try {
    $registryPath = "HKCU:\Software\AutoClicker\RecentFiles"
    
    if (!(Test-Path $registryPath)) {
      New-Item -Path $registryPath -Force | Out-Null
    }
    
    $recentFiles = Get-RecentFiles
    $recentFiles = @($FilePath) + ($recentFiles | Where-Object { $_ -ne $FilePath })
    
    if ($recentFiles.Count -gt 10) {
      $recentFiles = $recentFiles[0..9]
    }
    
    for ($i = 0; $i -lt $recentFiles.Count; $i++) {
      Set-ItemProperty -Path $registryPath -Name "File$($i + 1)" -Value $recentFiles[$i]
    }
    
    for ($i = $recentFiles.Count + 1; $i -le 10; $i++) {
      Remove-ItemProperty -Path $registryPath -Name "File$i" -ErrorAction SilentlyContinue
    }
    
  }
  catch {
    Write-Warning "Error updating recent files in registry: $($_.Exception.Message)"
  }
}

function Open-RecentFile {
  param([string]$FilePath)
  
  if (Test-Path $FilePath) {
    try {
      $jsonContent = Get-Content $FilePath -Raw
      $script:actions = ConvertFrom-Json $jsonContent
      $script:currentFile = $FilePath
      Add-RecentFile $FilePath
      Update-EventList
      $Global:statusLabel.Text = "Actions loaded from $FilePath"
    }
    catch {
      [System.Windows.Forms.MessageBox]::Show("Error loading file: $($_.Exception.Message)", "Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
  }
  else {
    [System.Windows.Forms.MessageBox]::Show("File not found: $FilePath", "File Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
  }
}

# Update lists
function Update-EventList {
  $playButton.Enabled = ($recordButton.Enabled -and $script:actions.Count -gt 0)
  $eventListView.Items.Clear()
  foreach ($action in $script:actions) {
    $item = New-Object System.Windows.Forms.ListViewItem(($eventListView.Items.Count + 1).ToString())
    if ($null -ne $action.Type) { $item.SubItems.Add($action.Type) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }
    if ($null -ne $action.X) { $item.SubItems.Add($action.X.ToString()) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }
    if ($null -ne $action.Y) { $item.SubItems.Add($action.Y.ToString()) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }
    if ($null -ne $action.Value) { $item.SubItems.Add($action.Value.ToString()) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }
    if ($null -ne $action.Duration) { $item.SubItems.Add($action.Duration.ToString()) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }
    if ($null -ne $action.Repeat) { $item.SubItems.Add($action.Repeat.ToString()) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }
    if ($null -ne $action.Comment) { $item.SubItems.Add($action.Comment) | Out-Null }
    else { $item.SubItems.Add("") | Out-Null }

    $eventListView.Items.Add($item) | Out-Null
  }

  Update-EventListColors
}

function Update-EventListColors {
  foreach ($Item in $eventListView.Items) {
    if ($Item.Index -eq $script:selectedIndex) {
      $Item.BackColor = [System.Drawing.Color]::Blue
      $Item.ForeColor = [System.Drawing.Color]::White
    }
    else {
      $Item.BackColor = [System.Drawing.Color]::White
      $Item.ForeColor = [System.Drawing.Color]::Black
    }
  }
}

##### Record #####
$script:keyboardHookProc = [Win32+LowLevelProc] {
  param($nCode, $wParam, $lParam)
    
  if ($nCode -ge 0 -and $script:isRecording) {
    $currentTime = [Environment]::TickCount
    $relativeTime = $currentTime - $script:prevTime
        
    try {
      # Read the keyboard data manually from memory
      $vkCode = [System.Runtime.InteropServices.Marshal]::ReadInt32($lParam)
      $scanCode = [System.Runtime.InteropServices.Marshal]::ReadInt32($lParam, 4)
      $action = if ($wParam -eq 0x0100) { "KeyDown" } else { "KeyUp" }
            
      # Check for duplicate keydown events
      $shouldRecord = $true
      if ($action -eq "KeyDown") {
        # If this key is already pressed (has a KeyDown without KeyUp), skip this duplicate
        if ($script:keyStates.ContainsKey($scanCode) -and $script:keyStates[$scanCode] -eq "KeyDown") {
          $shouldRecord = $false
        }
        else {
          # Mark this key as pressed
          $script:keyStates[$scanCode] = "KeyDown"
        }
      }
      elseif ($action -eq "KeyUp") {
        # Mark this key as released
        $script:keyStates[$scanCode] = "KeyUp"
      }
            
      # Only record if we should (no duplicate keydown)
      if ($shouldRecord) {
        $eventData = @{
          Type     = $action
          Duration = $relativeTime
          Value    = [System.Windows.Forms.Keys]$vkCode.ToString()
          scanCode = $scanCode
          Repeat   = 1
          Comment  = ""
        }

        $script:prevTime = $currentTime
        $script:actions += $eventData

        Update-EventList
      }
    }
    catch {
      # Ignore errors in hook
    }
  }
    
  return [Win32]::CallNextHookEx([IntPtr]::Zero, $nCode, $wParam, $lParam)
}

$script:mouseHookProc = [Win32+LowLevelProc] {
  param($nCode, $wParam, $lParam)
    
  if ($nCode -ge 0 -and $script:isRecording -and $script:targetHwnd) {
    $currentTime = [Environment]::TickCount
    $relativeTime = $currentTime - $script:prevTime
        
    try {
      # Read mouse data manually from memory
      $screenX = [System.Runtime.InteropServices.Marshal]::ReadInt32($lParam)
      $screenY = [System.Runtime.InteropServices.Marshal]::ReadInt32($lParam, 4)
            
      # Convert screen coordinates to window-relative coordinates
      $rect = New-Object Win32+RECT
      [Win32]::GetWindowRect($script:targetHwnd, [ref]$rect) | Out-Null
      $relX = $screenX - $rect.Left
      $relY = $screenY - $rect.Top
            
      # Only record if click is within the target window
      if ($relX -ge 0 -and $relY -ge 0 -and $relX -lt ($rect.Right - $rect.Left) -and $relY -lt ($rect.Bottom - $rect.Top)) {
        $action = switch ($wParam) {
          0x0201 { "Down" }    # WM_LBUTTONDOWN
          0x0202 { "Up" }      # WM_LBUTTONUP
          0x0204 { "Down" }    # WM_RBUTTONDOWN
          0x0205 { "Up" }      # WM_RBUTTONUP
          0x0207 { "Down" }    # WM_MBUTTONDOWN
          0x0208 { "Up" }      # WM_MBUTTONUP
          default { $null }
        }

        $value = switch ($wParam) {
          0x0201 { "Left" }    # WM_LBUTTONDOWN
          0x0202 { "Left" }    # WM_LBUTTONUP
          0x0204 { "Right" }   # WM_RBUTTONDOWN
          0x0205 { "Right" }   # WM_RBUTTONUP
          0x0207 { "Middle" }  # WM_MBUTTONDOWN
          0x0208 { "Middle" }  # WM_MBUTTONUP
          default { $null }
        }

        # wParam definitions for mouse messages:
        # 0x0201 = WM_LBUTTONDOWN  (Left mouse button down)
        # 0x0202 = WM_LBUTTONUP    (Left mouse button up)
        # 0x0204 = WM_RBUTTONDOWN  (Right mouse button down)
        # 0x0205 = WM_RBUTTONUP    (Right mouse button up)
        # 0x0207 = WM_MBUTTONDOWN  (Middle mouse button down)
        # 0x0208 = WM_MBUTTONUP    (Middle mouse button up)
        # 0x020A = WM_MOUSEWHEEL   (Mouse wheel scroll)
        # 0x020B = WM_XBUTTONDOWN  (X button down - side buttons)
        # 0x020C = WM_XBUTTONUP    (X button up - side buttons)
        # 0x0200 = WM_MOUSEMOVE    (Mouse movement)
                
        if ($action) {
          $eventData = @{
            Type     = "Mouse" + $action
            Value    = $value
            Duration = $relativeTime
            X        = $relX  # Window-relative coordinates
            Y        = $relY  # Window-relative coordinates
            Repeat   = 1
            Comment  = ""
          }
                    
          $script:prevTime = $currentTime
          $script:actions += $eventData

          Update-EventList
        }
      }
    }
    catch {
      # Ignore errors in hook
    }
  }
    
  return [Win32]::CallNextHookEx([IntPtr]::Zero, $nCode, $wParam, $lParam)
}
function Start-Recording {
  if (-not $script:targetHwnd) {
    $statusLabel.Text = "Status: Please select a target window first"
    return
  }
  $script:recordedEvents = @()
  $script:keyStates = @{}  # Clear key states at start of recording
  $script:prevTime = [Environment]::TickCount
  $script:isRecording = $true
    
  # Install keyboard hook
  $script:keyboardHook = [Win32]::SetWindowsHookEx(
    [Win32]::WH_KEYBOARD_LL,
    $script:keyboardHookProc,
    [Win32]::GetModuleHandle($null),
    0
  )
    
  # Install mouse hook
  $script:mouseHook = [Win32]::SetWindowsHookEx(
    [Win32]::WH_MOUSE_LL,
    $script:mouseHookProc,
    [Win32]::GetModuleHandle($null),
    0
  )
    
  $statusLabel.Text = "Status: Recording input relative to selected window..."
}

function Stop-Recording {
  $script:isRecording = $false
    
  if ($script:keyboardHook -ne [IntPtr]::Zero) {
    [Win32]::UnhookWindowsHookEx($script:keyboardHook) | Out-Null
    $script:keyboardHook = [IntPtr]::Zero
  }
    
  if ($script:mouseHook -ne [IntPtr]::Zero) {
    [Win32]::UnhookWindowsHookEx($script:mouseHook) | Out-Null
    $script:mouseHook = [IntPtr]::Zero
  }
    
  $statusLabel.Text = "Status: Recording stopped. Events recorded: $($script:recordedEvents.Count)"
    
  # Update events viewer if it's open
  if ($script:eventsViewerForm -and $script:eventsListBox) {
    Update-EventsList
  }
}

##### Playback #####
function Start-Playback {
  if (-not $script:targetHwnd) {
    $statusLabel.Text = "Status: Select a target window first"
    return
  }
    
  if ($script:actions.Count -eq 0) {
    $statusLabel.Text = "Status: No actions available. Record something or load a file first."
    return
  }
    
  # Get loop count from input
  try {
    $script:playbackLoopCount = [int]$loopCountBox.Text
    if ($script:playbackLoopCount -lt 0) {
      $script:playbackLoopCount = 1
      $loopCountBox.Text = "1"
    }
  }
  catch {
    $script:playbackLoopCount = 1
    $loopCountBox.Text = "1"
  }
    
  $script:isPlaying = $true
  $script:playbackIndex = 0
  $script:playbackCurrentLoop = 0
  $script:pressedKeys.Clear()  # Clear any previously tracked pressed keys
  $script:pressedMouseButtons.Clear()  # Clear any previously tracked pressed mouse buttons
  $script:prevTime = [Environment]::TickCount
  $script:playbackTimer.Interval = 10  # Check every 10ms
  $script:playbackTimer.Start()
    
  if ($script:playbackLoopCount -eq 0) {
    $statusLabel.Text = "Status: Playing back recording... (infinite loops)"
  }
  else {
    $statusLabel.Text = "Status: Playing back recording... (loop 1/$script:playbackLoopCount)"
  }
}

function Stop-Playback {
  $script:isPlaying = $false
  $script:playbackTimer.Stop()
    
  # Release all currently pressed keys
  if ($script:targetHwnd -and $script:pressedKeys.Count -gt 0) {
    foreach ($keyCode in $script:pressedKeys.Keys) {
      try {
        [Win32]::PostMessage($script:targetHwnd, $global:WM_KEYUP, [IntPtr]$keyCode, [IntPtr]0) | Out-Null
      }
      catch {
        # Ignore errors
      }
    }
    $script:pressedKeys.Clear()
  }
    
  # Release all currently pressed mouse buttons
  if ($script:targetHwnd -and $script:pressedMouseButtons.Count -gt 0) {
    # Get current window center for mouse release
    $rect = New-Object Win32+RECT
    [Win32]::GetWindowRect($script:targetHwnd, [ref]$rect) | Out-Null
    $centerX = ($rect.Right - $rect.Left) / 2
    $centerY = ($rect.Bottom - $rect.Top) / 2
    $lParam = [int]($centerY -shl 16) -bor [int]$centerX
        
    foreach ($button in $script:pressedMouseButtons.Keys) {
      try {
        switch ($button) {
          "Left" { 
            [Win32]::PostMessage($script:targetHwnd, [Win32]::WM_LBUTTONUP, [IntPtr]0, [IntPtr]$lParam) | Out-Null
          }
          "Right" { 
            [Win32]::PostMessage($script:targetHwnd, [Win32]::WM_RBUTTONUP, [IntPtr]0, [IntPtr]$lParam) | Out-Null
          }
        }
      }
      catch {
        # Ignore errors
      }
    }
    $script:pressedMouseButtons.Clear()
  }
    
  $statusLabel.Text = "Status: Playback stopped (all keys/buttons released)"
}
function Add-Action {
  param(
    [string]$Type,
    [int]$X,
    [int]$Y,
    [string]$Value,
    [int]$Duration,
    [int]$Repeat,
    [string]$Comment
  )
    
  $action = [PSCustomObject]@{
    X        = $null
    Y        = $null
    Value    = $null
    Type     = $Type
    Duration = $Duration
    Repeat   = $Repeat
    Comment  = $Comment
  }
  if ($X -ne $null) { $action.X = $X }
  if ($Y -ne $null) { $action.Y = $Y }
  if ($Value -ne $null) { $action.Value = $Value }

  $script:actions += $action
  Update-EventList
  $statusLabel.Text = "Status: Action added"
}
function Update-Action {
  param(
    [int]$Index,
    [string]$Type,
    [int]$X,
    [int]$Y,
    [string]$Value,
    [int]$Duration,
    [int]$Repeat,
    [string]$Comment
  )
    
  $action = [PSCustomObject]@{
    X        = $null
    Y        = $null
    Value    = $null
    Type     = $Type
    Duration = $Duration
    Repeat   = $Repeat
    Comment  = $Comment
  }
  if ($X -ne $null) { $action.X = $X }
  if ($Y -ne $null) { $action.Y = $Y }
  if ($Value -ne $null) { $action.Value = $Value }

  $script:actions[$Index] = $action
  Update-EventList
  $statusLabel.Text = "Status: Action updated"
}

function Remove-Action {
  param(
    [int]$Index
  )

  if ($Index -eq $null) {
    return
  }

  if ($Index -eq 0) {
    if ($script:actions.Count -gt 1) {
      $script:actions = $script:actions[1..($script:actions.Count - 1)]
    }
    else {
      $script:actions = @()
    }
  }
  elseif ($Index -eq ($script:actions.Count - 1)) {
    $script:actions = $script:actions[0..($script:actions.Count - 2)]
  }
  elseif ($Index -ge 0 -and $Index -lt $script:actions.Count) {
    $script:actions = $script:actions[0..($Index - 1)] + $script:actions[($Index + 1)..($script:actions.Count - 1)]
  }

  Update-EventList
  $statusLabel.Text = "Status: Action deleted"
}

function Move-Action {
  param($Index, $Direction)

  if ($Index -lt 0 -or $Index -ge $script:actions.Count) {
    return
  }

  $newIndex = $Index + $Direction
  if ($newIndex -lt 0 -or $newIndex -ge $script:actions.Count) {
    return
  }

  # Swap actions
  $script:actions[$Index], $script:actions[$newIndex] = $script:actions[$newIndex], $script:actions[$Index]
  Update-EventList
}

# Playback timer event
$script:playbackTimer.Add_Tick({
  # Check if we've finished all events in current loop
  if (-not $script:isPlaying -or $script:playbackIndex -ge $script:actions.Count) {
    # Check if we need to loop
    if ($script:isPlaying -and $script:actions.Count -gt 0) {
      $script:playbackCurrentLoop++
      
      # If loop count is 0 (infinite) or we haven't reached the loop limit
      if ($script:playbackLoopCount -eq 0 -or $script:playbackCurrentLoop -lt $script:playbackLoopCount) {
        # Reset for next loop
        $script:playbackIndex = 0
        $script:prevTime = [Environment]::TickCount
        
        # Update status
        if ($script:playbackLoopCount -eq 0) {
          $statusLabel.Text = "Status: Playing back recording... (loop $($script:playbackCurrentLoop + 1)/Infinity)"
        } else {
          $statusLabel.Text = "Status: Playing back recording... (loop $($script:playbackCurrentLoop + 1)/$script:playbackLoopCount)"
        }
        return
      }
    }
    
    # If we reach here, playback is complete
    $playButton.Text = "Play"
    $recordButton.Enabled = $True
    Stop-Playback
    return
  }
  
  $currentAction = $script:actions[$script:playbackIndex]
  $currentTime = [Environment]::TickCount
  
  # Check if it's time to execute this action
  if ($script:playbackIndex -eq 0 -or ($currentTime - $script:prevTime) -ge $currentAction.Duration) {  
    $script:prevTime = $currentTime
    try {
      # Handle keyboard actions
      if ($currentAction.Type -like "Key*") {
        $keyCode = [int][System.Windows.Forms.Keys]::($currentAction.Value)
        switch ($currentAction.Type) {
          "KeyDown" {
            Send-KeyDown $script:targetHwnd $keyCode | Out-Null
            $script:pressedKeys[$keyCode] = $true
          }
          "KeyUp" {
            Send-KeyUp $script:targetHwnd $keyCode | Out-Null
            $script:pressedKeys.Remove($keyCode)
          }
          "KeyPress" {
            Send-KeyPress $script:targetHwnd $keyCode | Out-Null
          }
        }
      }
      # Handle mouse actions
      elseif ($currentAction.Type -like "Mouse*") {
        $relX = [int]$currentAction.X
        $relY = [int]$currentAction.Y
        
        # Validate coordinates are within window bounds
        $rect = New-Object Win32+RECT
        [Win32]::GetWindowRect($script:targetHwnd, [ref]$rect) | Out-Null
        $windowWidth = $rect.Right - $rect.Left
        $windowHeight = $rect.Bottom - $rect.Top
        
        if ($relX -ge 0 -and $relY -ge 0 -and $relX -lt $windowWidth -and $relY -lt $windowHeight) {
          switch ($currentAction.Type) {
            "MouseDown" {
              Send-MouseDown $script:targetHwnd $relX $relY $currentAction.Value
              $script:pressedMouseButtons[$currentAction.Value] = $true
            }
            "MouseUp" {
              Send-MouseUp $script:targetHwnd $relX $relY $currentAction.Value
              $script:pressedMouseButtons[$currentAction.Value] = $false
            }
            "MouseClick" {
              Send-MouseClick $script:targetHwnd $relX $relY $currentAction.Value
            }
          }
        }
      }
    }
    catch {
      Write-Host "Playback error: $($_.Exception.Message)"
    }
    
    # Move to next action
    $script:playbackIndex++
    $script:playbackStartTime = $currentTime
  }
  
  # Update status with progress
  if ($script:playbackIndex -lt $script:actions.Count) {
    $progress = [math]::Round(($script:playbackIndex / $script:actions.Count) * 100)
    if ($script:playbackLoopCount -eq 0) {
      $statusLabel.Text = "Status: Playing back... $progress% (loop $($script:playbackCurrentLoop + 1)/Infinity)"
    } else {
      $statusLabel.Text = "Status: Playing back... $progress% (loop $($script:playbackCurrentLoop + 1)/$script:playbackLoopCount)"
    }
  }
})

$script:windowMap = @{}
function Update-WindowList {
  $windowList.Items.Clear()
  $script:windowMap.Clear()
  $callback = {
    param($hwnd, $lParam)
    if ([Win32]::IsWindowVisible($hwnd)) {
      $length = [Win32]::GetWindowTextLength($hwnd)
      if ($length -gt 0) {
        $sb = New-Object System.Text.StringBuilder($length + 1)
        [Win32]::GetWindowText($hwnd, $sb, $sb.Capacity) | Out-Null
        $windowTitle = $sb.ToString()
        $processId = 0
        [Win32]::GetWindowThreadProcessId($hwnd, [ref]$processId) | Out-Null
        $processName = (Get-Process -Id $processId -ErrorAction SilentlyContinue).ProcessName
        if ($processName) {
          $displayText = "$processId - $processName - $windowTitle"
          $windowList.Items.Add($displayText)
          $script:windowMap[$displayText] = $hwnd
        }
      }
    }
    return $true 
  }
  [Win32]::EnumWindows($callback, [System.IntPtr]::Zero)
  $Global:statusLabel.Text = "Status: Window list refreshed."
}
