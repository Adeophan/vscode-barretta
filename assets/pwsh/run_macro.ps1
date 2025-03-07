# This is a PowerShell script template.
# It contains five placeholders for runtime arguments.
# {{argument1}} - rootPath
# {{argument2}} - fileName
# {{argument3}} - callMethod
# {{argument4}} - paramsText
# {{argument5}} - vbaPassword (not directly used in this script, but kept for consistency and potential future use)
param(
  [string] $rootPath = "{{argument1}}",
  [string] $fileName = "{{argument2}}",
  [string] $callMethod = "{{argument3}}",
  [string] $paramsText = "{{argument4}}",
  [string] $vbaPassword = "{{argument5}}" # Kept for consistency, may be used in future enhancements
)
Write-Output "[barretta] Start processing : run_macro.ps1"

[String]$excelPath = "$rootPath/excel_file/$fileName"
Write-Output "[barretta] Excel FilePath : $excelPath"

# Add Windows API functions via Add-Type - put this at the beginning of your script for better organization
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Helper {
    [DllImport("user32.dll", SetLastError=true, CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
}
public class WindowHelper {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
Add-Type -AssemblyName System.Windows.Forms

function Find-WindowLike {
    param([string]$partialTitle)
    Get-Process | Where-Object {$_.MainWindowTitle -like "*$partialTitle*"} |
    Select-Object Id, MainWindowHandle, MainWindowTitle
}
function Release-ComObject {
    param([System.Object]$obj)
    if ($null -ne $obj) {
        try {
            [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) > $null
        } catch {
            Write-Output "Error releasing COM object: $_"
        }
    }
}

try {
  $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
  Write-Output "[barretta] Excel application is not running."
}

try {
  [System.Boolean]$endClose = $False

  if ($null -eq $excel) {
    $endClose = $True
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $True # Keep visible for macro execution feedback
    $book = $excel.Workbooks.Open($excelPath)
    Write-Output "[barretta] Launch a new Excel application. : Visible Mode"

  } else {
    Write-Output "[barretta] Use the activated Excel application. : Visible Mode"
    Write-Output "[barretta] Looking for Excel file: $fileName"

    # First check if it's already open as a regular workbook
    $book = $null
    foreach ($bookItem in $excel.Workbooks) {
      Write-Output "[barretta] Found workbook: $($bookItem.Name) at path: $($bookItem.FullName)"
      if ($bookItem.Name -eq "$fileName" -or $bookItem.FullName -eq "$excelPath") {
        $book = $bookItem
        Write-Output "[barretta] '$fileName' detected in launched Excel application."
        break
      }
    }

    # Normalize paths for reliable comparison
    $normalizedExcelPath = [System.IO.Path]::GetFullPath($excelPath)
    $fileName = [System.IO.Path]::GetFileName($normalizedExcelPath)
    $isAddin = $fileName -match "\.xlam$"

    # If not found as regular workbook and it's an add-in, check add-ins collection
    if ($null -eq $book -and $isAddin) {
      Write-Output "[barretta] This is an Excel add-in (.xlam). Checking add-in collection..."

      # Check all installed add-ins
      $addinFound = $false
      foreach ($addin in $excel.AddIns) {
        if ($addin.Installed) {
          $addinPath = $null
          try {
            # Try to get the full path of the add-in
            $addinPath = [System.IO.Path]::GetFullPath($addin.Path)
            Write-Output "[barretta] Found installed add-in: $($addin.Name) at path: $addinPath"
          }
          catch {
            Write-Output "[barretta] Could not resolve path for add-in $($addin.Name): $_"
          }

          # Compare normalized paths or filenames
          if (($addinPath -and $addinPath -eq $normalizedExcelPath) -or
              ($addin.Name -eq $fileName) -or
              ($addin.Name -eq [System.IO.Path]::GetFileNameWithoutExtension($fileName))) {

            Write-Output "[barretta] Add-in '$fileName' matches installed add-in: $($addin.Name)"
            $addinFound = $true

            # Try multiple approaches to get the workbook
            try {
              # Approach 1: Try by name
              $book = $excel.Workbooks.Item($addin.Name)
              Write-Output "[barretta] Successfully got add-in workbook by name: $($addin.Name)"
            }
            catch {
              Write-Output "[barretta] Could not get workbook by add-in name: $_"
              try {
                # Approach 2: Try by full name
                foreach ($wb in $excel.Workbooks) {
                  if ($wb.FullName -eq $normalizedExcelPath) {
                    $book = $wb
                    Write-Output "[barretta] Found add-in workbook by path matching"
                    break
                  }
                }
              }
              catch {
                Write-Output "[barretta] Error during workbook search by path: $_"
              }
            }

            if ($null -ne $book) {
              Write-Output "[barretta] Successfully retrieved add-in workbook"
              break
            }
          }
        }
      }

      # If add-in was found but workbook wasn't retrieved successfully
      if ($addinFound -and $null -eq $book) {
        Write-Output "[barretta] Add-in was found but workbook retrieval failed. Attempting direct open..."
        try {
          $book = $excel.Workbooks.Open($normalizedExcelPath)
          $book.IsAddin = $true  # Ensure it's treated as an add-in
          Write-Output "[barretta] Reopened add-in '$fileName' as workbook."
        }
        catch {
          Write-Output "[barretta] Failed to open add-in directly: $_"
        }
      }
    }

    # If book is still null, open the file directly
    if ($null -eq $book) {
      $excel = New-Object -ComObject Excel.Application
      $excel.Visible = $True # Keep visible for macro execution feedback
      $book = $excel.Workbooks.Open($excelPath)
      Write-Output "[barretta] New '$fileName' book opened in activated Excel application."
    }
  }

  Write-Output "[barretta] Running macro: '$callMethod $paramsText'"
  $excel.Run("$callMethod $paramsText")
  Write-Output "[barretta] Macro execution completed."


} catch {
  Write-Error "[barretta] An error has occurred. : $_"

} finally {
  try {
      # Release workbook
      if ($null -ne $book) { Release-ComObject $book }

      # Release Excel
      if ($null -ne $excel) { Release-ComObject $excel }
  } catch {
      Write-Output "Error during cleanup: $_"
  } finally {
    $book = $null
    $excel = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Write-Output "[barretta] Garbage collection was executed."
    Write-Output "[barretta] Finish processing : run_macro.ps1"
  }
}
