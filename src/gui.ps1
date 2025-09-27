Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global GUI Form
$Global:gui = New-Object System.Windows.Forms.Form
$Global:gui.Text = "Autoclicker Tool"
$Global:gui.Size = New-Object System.Drawing.Size(400, 400)
$Global:gui.StartPosition = "CenterScreen"
$Global:gui.FormBorderStyle = 'Sizable'
$Global:gui.MinimumSize = New-Object System.Drawing.Size(400, 400)

######## Menu Strip ########
$menuStrip = New-Object System.Windows.Forms.MenuStrip

#### File Menu ####
$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$menuStrip.Items.Add($fileMenu)

#### Edit Menu ####
$editMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$editMenu.Text = "Edit"
$menuStrip.Items.Add($editMenu)

#### About Menu ####
$aboutMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenu.Text = "About"
$menuStrip.Items.Add($aboutMenu)

# Menu items
$newMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$newMenuItem.Text = "New"
$fileMenu.DropDownItems.Add($newMenuItem)

$openMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$openMenuItem.Text = "Open"
$fileMenu.DropDownItems.Add($openMenuItem)

# Recent Files submenu
$recentFilesMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$recentFilesMenuItem.Text = "Recent Files"
$fileMenu.DropDownItems.Add($recentFilesMenuItem)

function Set-RecentFiles($recentFiles) {
  $recentFilesMenuItem.DropDownItems.Clear()
  foreach ($file in $recentFiles) {
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = $file
    # You can add a click event handler here if needed
    $recentFilesMenuItem.DropDownItems.Add($item)
  }
}

# Add click event handler for recent files menu items
$recentFilesMenuItem.Add_DropDownItemClicked({
  param($_sender, $e)
  $selectedFile = $e.ClickedItem.Text
  Open-RecentFile -FilePath $selectedFile
})

$saveMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$saveMenuItem.Text = "Save"
$fileMenu.DropDownItems.Add($saveMenuItem)

$fileMenu.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$fileMenu.DropDownItems.Add($exitMenuItem)

$Global:gui.Controls.Add($menuStrip)

# Set keyboard shortcuts for menu items
$newMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::N
$openMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::O
$saveMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::S
$exitMenuItem.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Q

# Enable shortcuts to be displayed in menu
$newMenuItem.ShowShortcutKeys = $true
$openMenuItem.ShowShortcutKeys = $true
$saveMenuItem.ShowShortcutKeys = $true
$exitMenuItem.ShowShortcutKeys = $true

$newMenuItem.Add_Click({
  New-AutomationScript
})
$openMenuItem.Add_Click({
  Open-AutomationScript
})
$saveMenuItem.Add_Click({
  Save-AutomationScript
})
$exitMenuItem.Add_Click({
  $Global:gui.Close()
  $Global:gui.Dispose()
})

# Add KeyDown event handler to the form for global shortcuts
$Global:gui.Add_KeyDown({
  param($_sender, $e)
  if ($e.Control) {
    switch ($e.KeyCode) {
      [System.Windows.Forms.Keys]::N {
        New-AutomationScript
      }
      [System.Windows.Forms.Keys]::O {
        Open-AutomationScript
      }
      [System.Windows.Forms.Keys]::Q {
        $Global:gui.Close()
        $Global:gui.Dispose()
      }
      [System.Windows.Forms.Keys]::S {
        Save-AutomationScript
      }
    }
  }
})

# Ensure the form can receive key events
$Global:gui.KeyPreview = $true

######## Main #########
# Window Selection
$targetWindowLabel = New-Object System.Windows.Forms.Label
$targetWindowLabel.Text = "Target Window:"
$targetWindowLabel.Location = New-Object System.Drawing.Point(10, 34)
$targetWindowLabel.Size = New-Object System.Drawing.Size(90, 25)
$Global:gui.Controls.Add($targetWindowLabel)

$windowList = New-Object System.Windows.Forms.ComboBox
$windowList.Location = New-Object System.Drawing.Point(100, 32)
$windowList.Size = New-Object System.Drawing.Size(210, 25)
$windowList.DropDownStyle = 'DropDownList'
$windowList.Anchor = 'Top,Left,Right'
$Global:gui.Controls.Add($windowList)
$windowList.Add_SelectedIndexChanged({
  if ($windowList.SelectedIndex -ge 0) {
    $selectedWindow = $windowList.SelectedItem
    $Global:statusLabel.Text = "Target window selected: $selectedWindow"
    
    $script:targetHwnd = $script:windowMap[$selectedWindow]
    if ($script:targetHwnd -eq $null -or $script:targetHwnd -eq [System.IntPtr]::Zero) {
      $recordButton.Enabled = $False
      $playButton.Enabled = $False
      $statusLabel.Text = "Status: Error! Window handle is invalid. Refresh list."
      return
    } else {
      $recordButton.Enabled = $True
      if ($recordButton.Enabled -and $script:actions.Count -gt 0) {
        $playButton.Enabled = $True
      }
    }
  }
})

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point(315, 30)
$refreshButton.Size = New-Object System.Drawing.Size(60, 25)
$refreshButton.Anchor = 'Top,Right'
$Global:gui.Controls.Add($refreshButton)
$refreshButton.Add_Click({
  Update-WindowList
})

##### Recording controls #####
$actionPanel = New-Object System.Windows.Forms.GroupBox
$actionPanel.Text = "Add / Edit Action"
$actionPanel.Location = New-Object System.Drawing.Point(10, 60)
$actionPanel.Size = New-Object System.Drawing.Size(365, 185)
$actionPanel.Anchor = 'Top,Left,Right'
$Global:gui.Controls.Add($actionPanel) | Out-Null

$recordButton = New-Object System.Windows.Forms.Button
$recordButton.Location = New-Object System.Drawing.Point(10, 20)
$recordButton.Size = New-Object System.Drawing.Size(60, 25)
$recordButton.Text = "Record"
$recordButton.Enabled = $False
$actionPanel.Controls.Add($recordButton) | Out-Null
$recordButton.Add_Click({
  if ($script:isRecording) {
    $recordButton.Text = "Record"
    Stop-Recording
    $playButton.Enabled = $True
  } else {
    $recordButton.Text = "Stop Rec"
    Start-Recording
    $playButton.Enabled = $False
  }
})

$playButton = New-Object System.Windows.Forms.Button
$playButton.Location = New-Object System.Drawing.Point(75, 20)
$playButton.Size = New-Object System.Drawing.Size(60, 25)
$playButton.Text = "Play"
$playButton.Enabled = $False
$actionPanel.Controls.Add($playButton) | Out-Null
$playButton.Add_Click({
  if ($script:isPlaying) {
    $playButton.Text = "Play"
    Stop-Playback
    $recordButton.Enabled = $True
  } else {
    $playButton.Text = "Stop"
    Start-Playback
    $recordButton.Enabled = $False
  }
})

# Loop count label and input
$loopLabel = New-Object System.Windows.Forms.Label
$loopLabel.Location = New-Object System.Drawing.Point(140, 26)
$loopLabel.Size = New-Object System.Drawing.Size(30, 25)
$loopLabel.Text = "Loop:"
$actionPanel.Controls.Add($loopLabel) | Out-Null

$loopCountBox = New-Object System.Windows.Forms.NumericUpDown
$loopCountBox.Location = New-Object System.Drawing.Point(170, 23)
$loopCountBox.Size = New-Object System.Drawing.Size(60, 25)
$loopCountBox.Minimum = 0
$loopCountBox.Maximum = 9999
$loopCountBox.Value = 1
$actionPanel.Controls.Add($loopCountBox) | Out-Null

# Create tooltip for loop count box
$loopToolTip = New-Object System.Windows.Forms.ToolTip
$loopToolTip.SetToolTip($loopCountBox, "Number of loops (0 = infinite)")

# Event Selection
$actionTypePanel = New-Object System.Windows.Forms.GroupBox
$actionTypePanel.Text = "Action Type"
$actionTypePanel.Location = New-Object System.Drawing.Point(10, 55)
$actionTypePanel.Size = New-Object System.Drawing.Size(90, 80)
$actionPanel.Controls.Add($actionTypePanel) | Out-Null

$mouseRadio = New-Object System.Windows.Forms.RadioButton
$mouseRadio.Text = "Mouse"
$mouseRadio.Location = New-Object System.Drawing.Point(10, 15)
$mouseRadio.Size = New-Object System.Drawing.Size(60, 20)
$mouseRadio.Checked = $true
$actionTypePanel.Controls.Add($mouseRadio) | Out-Null

$keyboardRadio = New-Object System.Windows.Forms.RadioButton
$keyboardRadio.Text = "Keyboard"
$keyboardRadio.Location = New-Object System.Drawing.Point(10, 35)
$keyboardRadio.Size = New-Object System.Drawing.Size(75, 20)
$actionTypePanel.Controls.Add($keyboardRadio) | Out-Null

$delayRadio = New-Object System.Windows.Forms.RadioButton
$delayRadio.Text = "Delay"
$delayRadio.Location = New-Object System.Drawing.Point(10, 55)
$delayRadio.Size = New-Object System.Drawing.Size(60, 20)
$actionTypePanel.Controls.Add($delayRadio) | Out-Null

# Click/Press, Down/Up
$eventTypePanel = New-Object System.Windows.Forms.GroupBox
$eventTypePanel.Text = "Event Type"
$eventTypePanel.Location = New-Object System.Drawing.Point(105, 55)
$eventTypePanel.Size = New-Object System.Drawing.Size(80, 80)
$actionPanel.Controls.Add($eventTypePanel) | Out-Null

$pressRadio = New-Object System.Windows.Forms.RadioButton
$pressRadio.Text = "Click"
$pressRadio.Location = New-Object System.Drawing.Point(10, 15)
$pressRadio.Size = New-Object System.Drawing.Size(60, 20)
$pressRadio.Checked = $true
$eventTypePanel.Controls.Add($pressRadio) | Out-Null
$mouseRadio.Add_CheckedChanged({
  if ($mouseRadio.Checked) {
    $pressRadio.Text = "Click"
    $mouseBtnTypePanel.Visible = $true
  } else {
    $mouseBtnTypePanel.Visible = $false
  }
})

$keyboardRadio.Add_CheckedChanged({
  if ($keyboardRadio.Checked) {
    $pressRadio.Text = "Press"
  }
})

$delayRadio.Add_CheckedChanged({
  if ($delayRadio.Checked) {
    $eventTypePanel.Visible = $false
  } else {
    $eventTypePanel.Visible = $true
  }
})

$downRadio = New-Object System.Windows.Forms.RadioButton
$downRadio.Text = "Down"
$downRadio.Location = New-Object System.Drawing.Point(10, 35)
$downRadio.Size = New-Object System.Drawing.Size(60, 20)
$eventTypePanel.Controls.Add($downRadio) | Out-Null

$upRadio = New-Object System.Windows.Forms.RadioButton
$upRadio.Text = "Up"
$upRadio.Location = New-Object System.Drawing.Point(10, 55)
$upRadio.Size = New-Object System.Drawing.Size(60, 20)
$eventTypePanel.Controls.Add($upRadio) | Out-Null

# Left, Middle, Right Mouse
$mouseBtnTypePanel = New-Object System.Windows.Forms.GroupBox
$mouseBtnTypePanel.Text = "Mouse Button Type"
$mouseBtnTypePanel.Location = New-Object System.Drawing.Point(190, 55)
$mouseBtnTypePanel.Size = New-Object System.Drawing.Size(80, 80)
$actionPanel.Controls.Add($mouseBtnTypePanel) | Out-Null

$leftRadio = New-Object System.Windows.Forms.RadioButton
$leftRadio.Text = "Left"
$leftRadio.Location = New-Object System.Drawing.Point(10, 15)
$leftRadio.Size = New-Object System.Drawing.Size(60, 20)
$leftRadio.Checked = $true
$mouseBtnTypePanel.Controls.Add($leftRadio) | Out-Null

$middleRadio = New-Object System.Windows.Forms.RadioButton
$middleRadio.Text = "Middle"
$middleRadio.Location = New-Object System.Drawing.Point(10, 35)
$middleRadio.Size = New-Object System.Drawing.Size(60, 20)
$mouseBtnTypePanel.Controls.Add($middleRadio) | Out-Null

$rightRadio = New-Object System.Windows.Forms.RadioButton
$rightRadio.Text = "Right"
$rightRadio.Location = New-Object System.Drawing.Point(10, 55)
$rightRadio.Size = New-Object System.Drawing.Size(60, 20)
$mouseBtnTypePanel.Controls.Add($rightRadio) | Out-Null

# Coordinates
$coordinatePanel = New-Object System.Windows.Forms.GroupBox
$coordinatePanel.Text = "Pos"
$coordinatePanel.Location = New-Object System.Drawing.Point(275, 55)
$coordinatePanel.Size = New-Object System.Drawing.Size(75, 80)
$actionPanel.Controls.Add($coordinatePanel) | Out-Null
$mouseRadio.Add_CheckedChanged({
  if ($mouseRadio.Checked) {
    $coordinatePanel.Visible = $true
  } else {
    $coordinatePanel.Visible = $false
  }
})

$xLabel = New-Object System.Windows.Forms.Label
$xLabel.Text = "X:"
$xLabel.Location = New-Object System.Drawing.Point(5, 20)
$xLabel.Size = New-Object System.Drawing.Size(15, 15)
$coordinatePanel.Controls.Add($xLabel) | Out-Null

$xBox = New-Object System.Windows.Forms.NumericUpDown
$xBox.Location = New-Object System.Drawing.Point(20, 18)
$xBox.Size = New-Object System.Drawing.Size(50, 20)
$xBox.Minimum = 0
$xBox.Maximum = 9999
$xBox.Value = 0
$coordinatePanel.Controls.Add($xBox) | Out-Null

$yLabel = New-Object System.Windows.Forms.Label
$yLabel.Text = "Y:"
$yLabel.Location = New-Object System.Drawing.Point(5, 45)
$yLabel.Size = New-Object System.Drawing.Size(15, 15)
$coordinatePanel.Controls.Add($yLabel) | Out-Null

$yBox = New-Object System.Windows.Forms.NumericUpDown
$yBox.Location = New-Object System.Drawing.Point(20, 43)
$yBox.Size = New-Object System.Drawing.Size(50, 20)
$yBox.Minimum = 0
$yBox.Maximum = 9999
$yBox.Value = 0
$coordinatePanel.Controls.Add($yBox) | Out-Null

# Keyboard Input
$keyInputPanel = New-Object System.Windows.Forms.GroupBox
$keyInputPanel.Text = "Keyboard Input"
$keyInputPanel.Location = New-Object System.Drawing.Point(190, 55)
$keyInputPanel.Size = New-Object System.Drawing.Size(160, 80)
$keyInputPanel.Visible = $false
$actionPanel.Controls.Add($keyInputPanel) | Out-Null

$keyTextBox = New-Object System.Windows.Forms.TextBox
$keyTextBox.Location = New-Object System.Drawing.Point(10, 40)
$keyTextBox.Size = New-Object System.Drawing.Size(140, 50)
$keyTextBox.Multiline = $false
$keyTextBox.ReadOnly = $true
$keyTextBox.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
$keyTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
$keyInputPanel.Controls.Add($keyTextBox) | Out-Null
# Manual input checkbox
$manualInputCheckBox = New-Object System.Windows.Forms.CheckBox
$manualInputCheckBox.Text = "Manual Input"
$manualInputCheckBox.Location = New-Object System.Drawing.Point(10, 20)
$manualInputCheckBox.Size = New-Object System.Drawing.Size(100, 20)
$manualInputCheckBox.Checked = $false
$keyInputPanel.Controls.Add($manualInputCheckBox) | Out-Null

# Adjust key textbox position to accommodate checkbox
$keyTextBox.Location = New-Object System.Drawing.Point(10, 45)
$keyTextBox.Size = New-Object System.Drawing.Size(140, 25)

$manualInputCheckBox.Add_CheckedChanged({
  if ($manualInputCheckBox.Checked) {
    $keyTextBox.ReadOnly = $false
    $keyTextBox.Cursor = [System.Windows.Forms.Cursors]::IBeam
  } else {
    $keyTextBox.ReadOnly = $true
    $keyTextBox.Cursor = [System.Windows.Forms.Cursors]::Hand
  }
})
$keyboardRadio.Add_CheckedChanged({
  if ($keyboardRadio.Checked) {
    $keyInputPanel.Visible = $true
    $mouseBtnTypePanel.Visible = $false
  } else {
    $keyInputPanel.Visible = $false
  }
})

$keyTextBox.Add_Click({
  if ($manualInputCheckBox.Checked -ne $true) {
    # Show a simple input dialog to capture key press
    $keyDialog = New-Object System.Windows.Forms.Form
    $keyDialog.Text = "Press a Key"
    $keyDialog.Size = New-Object System.Drawing.Size(300, 150)
    $keyDialog.StartPosition = "CenterParent"
    $keyDialog.KeyPreview = $true
    $keyDialog.ShowInTaskbar = $false
    $keyDialog.FormBorderStyle = 'FixedDialog'
    $keyDialog.MaximizeBox = $false
    $keyDialog.MinimizeBox = $false
  
    $keyLabel = New-Object System.Windows.Forms.Label
    $keyLabel.Text = "Press any key to capture it..."
    $keyLabel.Location = New-Object System.Drawing.Point(20, 20)
    $keyLabel.Size = New-Object System.Drawing.Size(250, 30)
    $keyLabel.TextAlign = 'MiddleCenter'
    $keyDialog.Controls.Add($keyLabel) | Out-Null
    
    $keyDialog.Add_KeyDown({
      param($_sender, $e)
      $scanCode = [System.Windows.Forms.Keys]::None
      try {
        # Convert KeyCode to ScanCode using MapVirtualKey
        $vkCode = [int]$e.KeyCode
        $scanCode = [Win32]::MapVirtualKey($vkCode, 0)
        if ($scanCode -ne 0) {
          $keyTextBox.Text = "SC_$scanCode"
        } else {
          $keyTextBox.Text = $e.KeyCode.ToString()
        }
      } catch {
        $keyTextBox.Text = $e.KeyCode.ToString()
      }
      $keyDialog.Close()
    })
  
    $keyDialog.ShowDialog($Global:gui) | Out-Null
  }
})

# Duration
$durationLabel = New-Object System.Windows.Forms.Label
$durationLabel.Text = "Duration (ms)"
$durationLabel.Location = New-Object System.Drawing.Point(10, 140)
$durationLabel.Size = New-Object System.Drawing.Size(80, 15)
$actionPanel.Controls.Add($durationLabel) | Out-Null

$durationBox = New-Object System.Windows.Forms.NumericUpDown
$durationBox.Location = New-Object System.Drawing.Point(10, 155)
$durationBox.Size = New-Object System.Drawing.Size(80, 25)
$durationBox.Minimum = 0
$durationBox.Maximum = 99999999
$durationBox.Value = 100
$actionPanel.Controls.Add($durationBox) | Out-Null

# Repeat
$repeatLabel = New-Object System.Windows.Forms.Label
$repeatLabel.Text = "Repeat"
$repeatLabel.Location = New-Object System.Drawing.Point(95, 140)
$repeatLabel.Size = New-Object System.Drawing.Size(50, 15)
$actionPanel.Controls.Add($repeatLabel) | Out-Null

$repeatBox = New-Object System.Windows.Forms.NumericUpDown
$repeatBox.Location = New-Object System.Drawing.Point(95, 155)
$repeatBox.Size = New-Object System.Drawing.Size(50, 25)
$repeatBox.Minimum = 1
$repeatBox.Maximum = 9999
$repeatBox.Value = 1
$actionPanel.Controls.Add($repeatBox) | Out-Null

# Add/Update/Delete buttons
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Location = New-Object System.Drawing.Point(160, 150)
$buttonPanel.Size = New-Object System.Drawing.Size(200, 30)
$actionPanel.Controls.Add($buttonPanel) | Out-Null

$addButton = New-Object System.Windows.Forms.Button
$addButton.Text = "+"
$addButton.Location = New-Object System.Drawing.Point(0, 0)
$addButton.Size = New-Object System.Drawing.Size(25, 25)
$buttonPanel.Controls.Add($addButton) | Out-Null
$addButton.Add_Click({
  if ($mouseRadio.Checked) {
    $x = $xBox.Value
    $y = $yBox.Value
  } else {
    $x = $null
    $y = $null
  }
  if ($mouseRadio.Checked) {
    $action = "Mouse"
  } elseif ($keyboardRadio.Checked) {
    if ($keyTextBox.Text -eq "" -and $keyboardRadio.Checked) {
      [System.Windows.Forms.MessageBox]::Show("Please enter a key or text for the keyboard action.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
      return
    }
    $action = "Key"
  } else {
    $action = "Delay"
  }
  if ($action -ne "Delay") {
    if ($pressRadio.Checked) {
      if ($mouseRadio.Checked) {
        $action += "Click"
      } else {
        $action += "Press"
      }
    } elseif ($downRadio.Checked) {
      $action += "Down"
    } elseif ($upRadio.Checked) {
      $action += "Up"
    }

    if ($mouseRadio.Checked) {
      if ($leftRadio.Checked) {
        $actionValue += "Left"
      } elseif ($middleRadio.Checked) {
        $actionValue += "Middle"
      } elseif ($rightRadio.Checked) {
        $actionValue += "Right"
      }
    } else {
      if ($keyboardRadio.Checked) {
        $actionValue = $keyTextBox.Text
      }
    }
  }
  $duration = $durationBox.Value
  $repeat = $repeatBox.Value
  Add-Action $action $x $y $actionValue $duration $repeat ""
})

$updateButton = New-Object System.Windows.Forms.Button
$updateButton.Text = "Update"
$updateButton.Location = New-Object System.Drawing.Point(30, 0)
$updateButton.Size = New-Object System.Drawing.Size(60, 25)
$updateButton.Enabled = $false
$buttonPanel.Controls.Add($updateButton) | Out-Null
$updateButton.Add_Click({
  if ($script:selectedIndex -ne $null) {
    if ($mouseRadio.Checked) {
      $x = $xBox.Value
      $y = $yBox.Value
    } else {
      $x = $null
      $y = $null
    }
    if ($mouseRadio.Checked) {
      $action = "Mouse"
    } elseif ($keyboardRadio.Checked) {
      if ($keyTextBox.Text -eq "" -and $keyboardRadio.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a key or text for the keyboard action.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
      }
      $action = "Key"
    } else {
      $action = "Delay"
    }
    if ($action -ne "Delay") {
      if ($pressRadio.Checked) {
        if ($mouseRadio.Checked) {
          $action += "Click"
        } else {
          $action += "Press"
        }
      } elseif ($downRadio.Checked) {
        $action += "Down"
      } elseif ($upRadio.Checked) {
        $action += "Up"
      }

      if ($mouseRadio.Checked) {
        if ($leftRadio.Checked) {
          $actionValue += "Left"
        } elseif ($middleRadio.Checked) {
          $actionValue += "Middle"
        } elseif ($rightRadio.Checked) {
          $actionValue += "Right"
        }
      } else {
        if ($keyboardRadio.Checked) {
          $actionValue = $keyTextBox.Text
        }
      }
    }
    $duration = $durationBox.Value
    $repeat = $repeatBox.Value
    Update-Action $script:selectedIndex $action $x $y $actionValue $duration $repeat ""
  }
})

$deleteButton = New-Object System.Windows.Forms.Button
$deleteButton.Text = "-"
$deleteButton.Location = New-Object System.Drawing.Point(95, 0)
$deleteButton.Size = New-Object System.Drawing.Size(25, 25)
$deleteButton.Enabled = $false
$buttonPanel.Controls.Add($deleteButton) | Out-Null
$deleteButton.Add_Click({
  if ($script:selectedIndex -ne $null) {
    $selectedItem = $eventListView.SelectedItems[0]
    $script:selectedIndex = $selectedItem.Index

    Remove-Action $script:selectedIndex
  }
})

$upButton = New-Object System.Windows.Forms.Button
$upButton.Text = "▲"
$upButton.Location = New-Object System.Drawing.Point(125, 0)
$upButton.Size = New-Object System.Drawing.Size(25, 25)
$upButton.Enabled = $false
$buttonPanel.Controls.Add($upButton) | Out-Null
$upButton.Add_Click({
  if ($script:selectedIndex -ne $null) {
    Move-Action $script:selectedIndex -1
    $selectedIndex -= 1
    Update-EventListColors
  }
})

$downButton = New-Object System.Windows.Forms.Button
$downButton.Text = "▼"
$downButton.Location = New-Object System.Drawing.Point(155, 0)
$downButton.Size = New-Object System.Drawing.Size(25, 25)
$downButton.Enabled = $false
$buttonPanel.Controls.Add($downButton) | Out-Null
$downButton.Add_Click({
  if ($script:selectedIndex -ne $null) {
    Move-Action $script:selectedIndex 1
    $selectedIndex += 1
    Update-EventListColors
  }
})

##### Event List #####
$eventListView = New-Object System.Windows.Forms.ListView
$eventListView.Location = New-Object System.Drawing.Point(0, 250)
$eventListView.Size = New-Object System.Drawing.Size(385, 90)
$eventListView.View = 'Details'
$eventListView.FullRowSelect = $true
$eventListView.GridLines = $true
$eventListView.MultiSelect = $false
$eventListView.HideSelection = $true


$eventListView.Anchor = 'Top,Bottom,Left,Right'
# Add columns
$eventListView.Columns.Add("No", 30) | Out-Null
$eventListView.Columns.Add("Action", 80) | Out-Null
$eventListView.Columns.Add("X", 30) | Out-Null
$eventListView.Columns.Add("Y", 30) | Out-Null
$eventListView.Columns.Add("Key/Text", 80) | Out-Null
$eventListView.Columns.Add("Duration", 60) | Out-Null
$eventListView.Columns.Add("Repeat", 60) | Out-Null
$eventListView.Columns.Add("Comment", 120) | Out-Null
$Global:gui.Controls.Add($eventListView) | Out-Null

$eventListView.Add_SelectedIndexChanged({
  if ($eventListView.SelectedItems.Count -gt 0) {
    $selectedItem = $eventListView.SelectedItems[0]
    $script:selectedIndex = $selectedItem.Index
  }

  # Disable buttons when no item is selected
  if ($script:selectedIndex -eq $null) {
    $updateButton.Enabled = $false
    $deleteButton.Enabled = $false
    $upButton.Enabled = $false
    $downButton.Enabled = $false
  } else {
    # Enable update and delete buttons
    $updateButton.Enabled = $true
    $deleteButton.Enabled = $true
    $upButton.Enabled = ($script:selectedIndex -gt 0)
    $downButton.Enabled = ($script:selectedIndex -lt ($eventListView.Items.Count - 1))
  }

  Update-EventListColors
  
  # Populate the form fields with selected action data
  $action = $script:actions[$script:selectedIndex]
  
  # Set action type radio buttons
  $mouseRadio.Checked = $False
  $keyboardRadio.Checked = $False
  $delayRadio.Checked = $False
  if ($action.Type -like "Mouse*") {
    $mouseRadio.Checked = $true
    $xBox.Value = $action.X
    $yBox.Value = $action.Y
  } elseif ($action.Type -like "Key*") {
    $keyboardRadio.Checked = $true
    $keyTextBox.Text = $action.Value
  } elseif ($action.Type -eq "Delay") {
    $delayRadio.Checked = $true
  }
  
  # Set event type radio buttons
  $pressRadio.Checked = $False
  $downRadio.Checked = $False
  $upRadio.Checked = $False
  if ($action.Type -like "*Click" -or $action.Type -like "*Press") {
    $pressRadio.Checked = $true
  } elseif ($action.Type -like "*Down") {
    $downRadio.Checked = $true
  } elseif ($action.Type -like "*Up") {
    $upRadio.Checked = $true
  }
  
  # Set mouse button type for mouse actions
  $leftRadio.Checked = $False
  $middleRadio.Checked = $False
  $rightRadio.Checked = $False
  if ($action.Type -like "Mouse*") {
    if ($action.Value -eq "Left") {
      $leftRadio.Checked = $true
    } elseif ($action.Value -eq "Middle") {
      $middleRadio.Checked = $true
    } elseif ($action.Value -eq "Right") {
      $rightRadio.Checked = $true
    }
  } elseif ($action.Type -like "Key*") {
    $keyTextBox.Text = $action.Value
  }
  
  $durationBox.Value = $action.Duration
  $repeatBox.Value = $action.Repeat
  
  $Global:statusLabel.Text = "Action selected: $($action.Action) at position $($script:selectedIndex + 1)"
})

######## Status Bar ########
$statusBar = New-Object System.Windows.Forms.StatusStrip

$Global:statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$Global:statusLabel.Text = "Ready - Select target window and add actions"
$statusBar.Items.Add($Global:statusLabel) | Out-Null
$Global:gui.Controls.Add($statusBar) | Out-Null
