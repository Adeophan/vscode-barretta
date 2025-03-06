# This is a PowerShell script template.
# It contains four placeholders for runtime arguments.
# {{argument1}} - rootPath
# {{argument2}} - fileName
# {{argument3}} - pullIgnoreDocument
# {{argument4}} - vbaPassword
param(
  [string] $rootPath = "{{argument1}}",
  [string] $fileName = "{{argument2}}",
  [boolean] $pullIgnoreDocument = [boolean]"{{argument3}}",
  [string] $vbaPassword = "{{argument4}}"
)
Write-Output "[barretta] Start processing : pull_modules.ps1"

[String]$excelPath = "$rootPath/excel_file/$fileName"
Write-Output "[barretta] Excel FilePath : $excelPath"

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
      
      # Set IsAddin property if it's an add-in file
      if ($isAddin) {
        $book.IsAddin = $false
        Write-Output "[barretta] Add-in '$fileName' opened as regular workbook."
      } else {
        Write-Output "[barretta] New '$fileName' book opened in activated Excel application."
      }
    }
  }
  
  # Unlock the VBA project if a password is supplied.
  if ($vbaPassword -ne "") {
      Write-Output "[barretta] VBA project is password protected. Attempting to unlock via API..."
      try {
          # Ensure the VBA IDE is visible.
          $excel.VBE.MainWindow.Visible = $True
          . "$rootPath/barretta-core/scripts/unlock_vba_project.ps1"
          if (-not (Unlock-VBAProject $vbaPassword)) {
              Write-Output "Error: VBA project unlocking failed."
          } else {
              Write-Output "[barretta] VBA project unlocked successfully."
          }
      } catch {
          Write-Output "Error: Exception during VBA unlock: $_"
      }
  }
  
  foreach ($module in $book.VBProject.VBComponents) {
      # Module export logic remains unchanged.
      switch ($module.Type) {
          1 {
              $fileExtension = ".bas"
              $outPath = "$rootPath/code_modules/" + $module.Name + $fileExtension
              $module.Export($outPath)
              Write-Output "[barretta] Output : $outPath"
          }
          2 {
              $fileExtension = ".cls"
              $outPath = "$rootPath/code_modules/" + $module.Name + $fileExtension
              $module.Export($outPath)
              Write-Output "[barretta] Output : $outPath"
          }
          3 {
              $fileExtension = ".frm"
              $outPath = "$rootPath/code_modules/" + $module.Name + $fileExtension
              $module.Export($outPath)
              Write-Output "[barretta] Output : $outPath"
          }
          11 {
              $fileExtension = ".unknown"
              $outPath = "$rootPath/code_modules/" + $module.Name + $fileExtension
              $module.Export($outPath)
              Write-Output "[barretta] Output : $outPath"
          }
          100 {
              if (-not $pullIgnoreDocument) {
                  $fileExtension = ".cls"
                  $outPath = "$rootPath/code_modules/" + $module.Name + $fileExtension
                  $module.Export($outPath)
                  Write-Output "[barretta] Output : $outPath"
              }
          }
      }
  }
}
catch {
  Write-Error "[barretta] An error has occurred: $_"
}
finally {
  if ($endClose) {
    $excel.Quit()
    Write-Output "[barretta] Quit Excel"
  }
  [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) > $null
  $module = $null
  $book = $null
  $excel = $null
  [System.GC]::Collect()
  Write-Output "[barretta] Garbage collection was executed."
  Write-Output "[barretta] Finish processing : pull_modules.ps1"
}
