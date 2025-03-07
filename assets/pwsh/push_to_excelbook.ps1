# This is a PowerShell script template.
# It contains four placeholders for runtime arguments.
# {{argument1}} - rootPath
# {{argument2}} - fileName
# {{argument3}} - pushIgnoreDocument
# {{argument4}} - vbaPassword
param(
  [string] $rootPath = "{{argument1}}",
  [string] $fileName = "{{argument2}}",
  [boolean] $pushIgnoreDocument = [boolean]"{{argument3}}",
  [string] $vbaPassword = "{{argument4}}"
)
Write-Output "[barretta] Start processing : push_modules.ps1"

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
    $excel.Visible = $False
    $book = $excel.Workbooks.Open($excelPath)
    Write-Output "[barretta] Launch a new Excel application. : Invisible Mode"
  }
  else {
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
      $excel.Visible = $True
      $book = $excel.Workbooks.Open($excelPath)
      Write-Output "[barretta] New '$fileName' book opened in activated Excel application."
    }
  }

  # Only unlock the VBA project if it's actually password protected
  if ($vbaPassword -ne "") {
      $protection = $book.VBProject.Protection
      if ($protection -eq 1) {
          Write-Output "[barretta] VBA project is password protected ($protection). Attempting to unlock via API..."
          try {
              $vbe = $excel.VBE
              # Ensure the VBA IDE is visible
              $vbe.MainWindow.Visible = $True
              # Click on Specific VBA Project inside VBA IDE to get password prompt
              foreach ($project in $vbe.VBProjects) {
                Write-Output "[barretta] Project : $($project.FileName)"
                if ($project.FileName -eq $book.FullName) {
                  $vbe.ActiveVBProject = $project
                  Start-Sleep -Seconds 1
                  # Instead of Activate(), try to trigger password prompt by sending Enter to VBA window
                  Write-Output "[barretta] Attempting to send Enter key to VBA window to trigger password prompt."
                  $vbaMainWindowTitle = $vbe.MainWindow.Caption
                  $vbaWindow = Find-WindowLike $vbaMainWindowTitle
                  $vbaWindowHandle = $vbaWindow[0].MainWindowHandle

                  if ($vbaWindowHandle -ne [IntPtr]::Zero) {
                      Write-Output "[barretta] VBA window found $($vbaWindowHandle). Sending Enter key."
                      [WindowHelper]::SetForegroundWindow($vbaWindowHandle)
                      Start-Sleep -Milliseconds 100
                      [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                      Start-Sleep -Milliseconds 100
                  } else {
                      Write-Output "[barretta] VBA window not found. Password prompt may not be triggered."
                  }

                  Write-Output "[barretta] Caption: $($vbe.MainWindow.Caption)"
                  Write-Output "[barretta] Selected project : $($book.FullName)"
                  break
                }
              }

              . "$rootPath/barretta-core/scripts/unlock_vba_project.ps1"
              if (-not (Unlock-VBAProject $vbaPassword)) {
                  Write-Output "Error: VBA project unlocking failed."
              } else {
                  Write-Output "[barretta] VBA project unlocked successfully."
              }
          } catch {
              Write-Output "Error: Exception during VBA unlock: $_"
          }
      } else {
          Write-Output "[barretta] VBA project is not password protected ($protection). Skipping unlock."
      }
  }

  $distPath = "$rootPath/barretta-core/dist"
  Write-Output "[barretta] The target 'dist' path was set like this: $distPath"
  $moduleFiles = Get-ChildItem $distPath -File | Where-Object { $_.Name -match '.bas$|.cls$|.frm$' }

  Foreach ($moduleFile in $moduleFiles) {
    Write-Output "[barretta] Start file processing: $moduleFile"
    $moduleName = $moduleFile.BaseName
    $existModule = $False
    foreach ($module in $book.VBProject.VBComponents) {
      if ($module.Name -eq $moduleName) {
        $existModule = $True
      }
    }
    if ($existModule) {
      Write-Output "[barretta] '$moduleFile' exists in the Excel file and will be replaced."
      $module = $book.VBProject.VBComponents.Item($moduleName)
      switch ($module.Type) {
        1 {
          $trgModule = $book.VBProject.VBComponents.Item($moduleName)
          $book.VBProject.VBComponents.Remove($trgModule)
          Write-Output "[barretta] Module removed: $moduleName"
          Start-Sleep -Seconds 1
          $book.VBProject.VBComponents.Import($moduleFile.FullName)
          Write-Output "[barretta] Module imported: $moduleName"
        }
        2 {
          $trgModule = $book.VBProject.VBComponents.Item($moduleName)
          $book.VBProject.VBComponents.Remove($trgModule)
          Write-Output "[barretta] Module removed: $moduleName"
          Start-Sleep -Seconds 1
          $book.VBProject.VBComponents.Import($moduleFile.FullName)
          Write-Output "[barretta] Module imported: $moduleName"
        }
        3 {
          $trgModule = $book.VBProject.VBComponents.Item($moduleName)
          $book.VBProject.VBComponents.Remove($trgModule)
          Write-Output "[barretta] Module removed: $moduleName"
          Start-Sleep -Seconds 1
          $book.VBProject.VBComponents.Import($moduleFile.FullName)
          Write-Output "[barretta] Module imported: $moduleName"
        }
        11 {
          Write-Output "[barretta] FileType: ActiveX - no action taken."
        }
        100 {
          if (-not $pushIgnoreDocument) {
            Write-Output "[barretta] FileType: Document - action skipped."
          }
        }
      }
    }
    else {
      Write-Output "[barretta] '$moduleFile' is a new module; importing."
      $book.VBProject.VBComponents.Import($moduleFile.FullName)
    }
  }
}
catch {
  Write-Error "[barretta] An error has occurred: $_"
}
finally {
  try {
    if ($endClose) {
      $book.Save()
      Write-Output "[barretta] Excel file saved: $fileName"
      $book.Close()
      Write-Output "[barretta] Excel file closed: $fileName"
      $excel.Quit()
      Write-Output "[barretta] Excel application terminated."
    }
    else {
      $book.Save()
      Write-Output "[barretta] Excel file saved: $fileName"
    }
  }
  catch {
    Write-Warning "[barretta] Error during book save/close/quit: $_"
  }
  finally {
    try {
      if ($null -ne $vbe) { Release-ComObject $vbe }

      # Release each VBComponent
      if ($null -ne $book -and $null -ne $book.VBProject -and $null -ne $book.VBProject.VBComponents) {
          foreach ($component in $book.VBProject.VBComponents) {
              Release-ComObject $component
          }
      }

      # Release VBProject
      if ($null -ne $book -and $null -ne $book.VBProject) {
          Release-ComObject $book.VBProject
      }

      # Release workbook
      if ($null -ne $book) { Release-ComObject $book }

      # Close Excel only if we opened it
      if ($endClose -and $null -ne $excel) {
          $excel.Quit()
      }

      # Release Excel
      if ($null -ne $excel) { Release-ComObject $excel }
    } catch {
        Write-Output "Error during cleanup: $_"
    } finally {
      $vbe = $null
      $module = $null
      $book = $null
      $excel = $null
      [System.GC]::Collect()
      [System.GC]::WaitForPendingFinalizers()
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect()
    Write-Output "[barretta] Garbage collection was executed."
    Write-Output "[barretta] Finish processing : push_modules.ps1"
  }
}


