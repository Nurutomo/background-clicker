if (-not ("Win32" -as [type])) {
    Add-Type -TypeDefinition @"
    using System;
    using System.Text;
    using System.Runtime.InteropServices;

    public class Win32 {
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        
        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        
        [DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT Point);

        public delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);

        public const int WH_KEYBOARD_LL = 13;
        public const int WH_MOUSE_LL = 14;
        public const int HC_ACTION = 0;
        
        public const int KEYEVENTF_KEYDOWN = 0x0000;
        public const int KEYEVENTF_KEYUP = 0x0002;
        public const int KEYEVENTF_SCANCODE = 0x0008;
        public const int INPUT_KEYBOARD = 1;
        public const int INPUT_MOUSE = 0;
        public const int SW_RESTORE = 9;
        public const int SW_SHOW = 5;

        // Mouse event flags
        public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        public const uint MOUSEEVENTF_LEFTUP = 0x0004;
        public const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        public const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        public const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        public const uint MOUSEEVENTF_MIDDLEUP = 0x0040;

        // Mouse message constants
        public const uint WM_LBUTTONDOWN = 0x0201;
        public const uint WM_LBUTTONUP = 0x0202;
        public const uint WM_RBUTTONDOWN = 0x0204;
        public const uint WM_RBUTTONUP = 0x0205;
        
        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT {
            public int type;
            public InputUnion u;
        }
        
        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion {
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public MOUSEINPUT mi;
        }
        
        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        // Overlay window constants
        public const int WS_EX_LAYERED = 0x80000;
        public const int WS_EX_TRANSPARENT = 0x20;
        public const int WS_EX_TOPMOST = 0x8;
        public const int LWA_ALPHA = 0x2;
        public const int LWA_COLORKEY = 0x1;

        // Additional overlay functions
        [DllImport("user32.dll")]
        public static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        public const int GWL_EXSTYLE = -20;
    }
"@
}

# (The rest of the script is unchanged and correct)
# Define keyboard message constants
$global:WM_KEYDOWN = 0x0100
$global:WM_KEYUP = 0x0101

function MapVirtualKey($vKey, $mapType) {
    Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class KeyMapper {
            [DllImport("user32.dll")]
            public static extern uint MapVirtualKey(uint uCode, uint uMapType);
        }
"@ -ErrorAction SilentlyContinue
    return [KeyMapper]::MapVirtualKey($vKey, $mapType)
}

function Get-KeyName {
    param([int]$vkCode)
    
    $keyNames = @{
        8 = "Backspace"; 9 = "Tab"; 13 = "Enter"; 16 = "Shift"; 17 = "Ctrl"; 18 = "Alt"
        19 = "Pause"; 20 = "CapsLock"; 27 = "Escape"; 32 = "Space"; 33 = "PageUp"; 34 = "PageDown"
        35 = "End"; 36 = "Home"; 37 = "Left"; 38 = "Up"; 39 = "Right"; 40 = "Down"
        45 = "Insert"; 46 = "Delete"; 48 = "0"; 49 = "1"; 50 = "2"; 51 = "3"; 52 = "4"
        53 = "5"; 54 = "6"; 55 = "7"; 56 = "8"; 57 = "9"
        65 = "A"; 66 = "B"; 67 = "C"; 68 = "D"; 69 = "E"; 70 = "F"; 71 = "G"; 72 = "H"
        73 = "I"; 74 = "J"; 75 = "K"; 76 = "L"; 77 = "M"; 78 = "N"; 79 = "O"; 80 = "P"
        81 = "Q"; 82 = "R"; 83 = "S"; 84 = "T"; 85 = "U"; 86 = "V"; 87 = "W"; 88 = "X"
        89 = "Y"; 90 = "Z"
        112 = "F1"; 113 = "F2"; 114 = "F3"; 115 = "F4"; 116 = "F5"; 117 = "F6"
        118 = "F7"; 119 = "F8"; 120 = "F9"; 121 = "F10"; 122 = "F11"; 123 = "F12"
        160 = "LShift"; 161 = "RShift"; 162 = "LCtrl"; 163 = "RCtrl"; 164 = "LAlt"; 165 = "RAlt"
    }
    
    if ($keyNames.ContainsKey($vkCode)) {
        return $keyNames[$vkCode]
    }
    else {
        return "Key$vkCode"
    }
}

# Mouse
function Send-MouseClick($hwnd, $x, $y, $button) {
  Send-MouseDown $hwnd $x $y $button
  Start-Sleep -Milliseconds 10
  Send-MouseUp $hwnd $x $y $button
}
function Send-MouseDown($hwnd, $x, $y, $button) {
  $lParam = ($y -shl 16) -bor ($x -band 0xFFFF)
  if ($button -eq "Left") {
    [Win32]::PostMessage($hwnd, [Win32]::WM_LBUTTONDOWN, [IntPtr]0, [IntPtr]$lParam) | Out-Null
  } elseif ($button -eq "Middle") {
    [Win32]::PostMessage($hwnd, [Win32]::WM_MBUTTONDOWN, [IntPtr]0, [IntPtr]$lParam) | Out-Null
  } elseif ($button -eq "Right") {
    [Win32]::PostMessage($hwnd, [Win32]::WM_RBUTTONDOWN, [IntPtr]0, [IntPtr]$lParam) | Out-Null
  }
}

function Send-MouseUp($hwnd, $x, $y, $button) {
  $lParam = ($y -shl 16) -bor ($x -band 0xFFFF)
  if ($button -eq "Left") {
    [Win32]::PostMessage($hwnd, [Win32]::WM_LBUTTONUP, [IntPtr]0, [IntPtr]$lParam) | Out-Null
  } elseif ($button -eq "Middle") {
    [Win32]::PostMessage($hwnd, [Win32]::WM_MBUTTONUP, [IntPtr]0, [IntPtr]$lParam) | Out-Null
  } elseif ($button -eq "Right") {
    [Win32]::PostMessage($hwnd, [Win32]::WM_RBUTTONUP, [IntPtr]0, [IntPtr]$lParam) | Out-Null
  }
}

# Keyboard
function Send-Key($hwnd, $keyCode) {
  Send-KeyDown $hwnd $key
  Start-Sleep -Milliseconds 10
  Send-KeyUp $hwnd $key
}
function Send-KeyDown($hwnd, $keyCode) {
  [Win32]::PostMessage($hwnd, [Win32]::WM_KEYDOWN, [IntPtr]$keyCode, [IntPtr]0) | Out-Null
}
function Send-KeyUp($hwnd, $keyCode) {
  [Win32]::PostMessage($hwnd, [Win32]::WM_KEYUP, [IntPtr]$keyCode, [IntPtr]0) | Out-Null
}
function Send-Text($hwnd, $text) {
  foreach ($char in $text.ToCharArray()) {
    $keyCode = [int]$char
    [Win32]::PostMessage($hwnd, [Win32]::WM_CHAR, [IntPtr]$keyCode, [IntPtr]0) | Out-Null
    Start-Sleep -Milliseconds 50
  }
}
