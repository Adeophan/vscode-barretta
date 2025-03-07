# Add Windows API functions via Add-Type
Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class Win32 {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam);

    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildCallback lpEnumFunc, IntPtr lParam);
    
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern uint GetDlgCtrlID(IntPtr hwnd);
    
    public delegate bool EnumChildCallback(IntPtr hwnd, IntPtr lParam);
}
"@

function Unlock-VBAProject {
    param(
        [string]$vbaPassword
    )

    # window title and API constants
    $vbaWindowTitle = "VBAProject Password"
    $WM_SETTEXT = 0xC
    $BM_CLICK = 0xF5

    # Attempt to locate the password prompt window
    $vbaWinHandle = [Win32]::FindWindow("#32770", $vbaWindowTitle)
    Write-Output "Debug: VBA window handle returned: $vbaWinHandle"
    if ($vbaWinHandle -eq [IntPtr]::Zero) {
        Write-Output "Error: VBAProject Password window was not found."
        return $false
    }

    # Find the Edit control (text box) within the VBA window
    $editHandle = [Win32]::FindWindowEx($vbaWinHandle, [IntPtr]::Zero, "Edit", $null)
    if ($editHandle -eq [IntPtr]::Zero) {
        Write-Output "Error: Password input control was not found."
        return $false
    }

    # Send the password directly to the Edit control
    [Win32]::SendMessage($editHandle, $WM_SETTEXT, [IntPtr]::Zero, $vbaPassword) | Out-Null
    Start-Sleep -Milliseconds 100

    # Initialize our button collection outside the callback
    $script:buttonList = New-Object System.Collections.ArrayList
    
    # We'll enumerate all child windows and find all buttons
    $enumCallback = {
        param([IntPtr]$hwnd, [IntPtr]$lParam)
        
        $classNameBuilder = New-Object System.Text.StringBuilder 256
        [Win32]::GetClassName($hwnd, $classNameBuilder, 256)
        $className = $classNameBuilder.ToString()
        
        $windowTextBuilder = New-Object System.Text.StringBuilder 256
        [Win32]::GetWindowText($hwnd, $windowTextBuilder, 256)
        $windowText = $windowTextBuilder.ToString()
        
        # If this is a button, add it to our collection
        if ($className -eq "Button") {
            $ctrlId = [Win32]::GetDlgCtrlID($hwnd)
            Write-Output "Found button: Handle=$hwnd, Text='$windowText', ControlID=$ctrlId"
            
            # Create button object
            $buttonObj = [PSCustomObject]@{
                Handle = $hwnd
                Text = $windowText
                ControlID = $ctrlId
            }
            
            # Add to ArrayList (safer than using +=)
            [void]$script:buttonList.Add($buttonObj)
        }
        
        return $true  # Continue enumeration
    }
    
    $callbackDelegate = [Win32+EnumChildCallback]$enumCallback
    [Win32]::EnumChildWindows($vbaWinHandle, $callbackDelegate, [IntPtr]::Zero)
    
    Write-Output "Total buttons found: $($script:buttonList.Count)"
    
    # Button selection logic - try to find the best match for "OK" button
    $buttonToClick = $null
    
    # Try multiple approaches to find the OK button
    # First, look for a button with "OK" text
    foreach ($button in $script:buttonList) {
        if ($button.Text -eq "OK") {
            $buttonToClick = $button
            Write-Output "Found button with 'OK' text"
            break
        }
    }
    
    # If not found, look for common OK button control IDs (1 or 2 are common for OK buttons)
    if (-not $buttonToClick) {
        foreach ($button in $script:buttonList) {
            if ($button.ControlID -eq 1 -or $button.ControlID -eq 2) {
                $buttonToClick = $button
                Write-Output "Found button with standard OK ControlID"
                break
            }
        }
    }
    
    # If still not found and we have buttons, use the first one
    if (-not $buttonToClick -and $script:buttonList.Count -gt 0) {
        $buttonToClick = $script:buttonList[0]
        Write-Output "Using first button found as fallback"
    }
    
    # Click the button if found
    if ($buttonToClick) {
        Write-Output "Clicking button: Handle=$($buttonToClick.Handle), Text='$($buttonToClick.Text)', ControlID=$($buttonToClick.ControlID)"
        [Win32]::SendMessage($buttonToClick.Handle, $BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
        Write-Output "VBA project unlocked via API call."
        return $true
    } else {
        Write-Output "Error: OK button not found in the VBAProject Password window."
        return $false
    }
}
