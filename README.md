# Barretta for Visual Studio Code

<img src="https://github.com/Mikoshiba-Kyu/vscode-barretta/blob/main/docs/image/largeicon_750x256.png?raw=true" width="450px">

## What is Barretta?

Barretta is an extension for Excel VBA developers.

<img src="https://github.com/Mikoshiba-Kyu/vscode-barretta/blob/main/docs/image/commands.gif?raw=true">

<img src="https://github.com/Mikoshiba-Kyu/vscode-barretta/blob/main/docs/image/launcher.gif?raw=true">

It serves as a coding assistant in Visual Studio Code with the ability to interact with Excel files via various commands.  
It is analogous to clasp in Google Apps Script (used to fasten) and serves as a "clip" for Excel VBA.

## Quick Start

1. Open an empty folder in Visual Studio Code and run the `Barretta: Init` command.
2. Place the Excel file (macro-enabled) you wish to edit into the **excel_file** folder.
   > For the initial run, it is recommended to use an Excel file that already contains at least one standard module.
3. Execute the `Barretta: Pull` command.
4. Edit the files exported to the **code_modules** folder in Visual Studio Code.
5. Import the updated modules into the Excel file using the `Barretta: Push` command.

## Commands

The following commands can be executed from the Visual Studio Code command palette (`Ctrl + Shift + P`):

- **Barretta: Init** (Alt + Ctrl + N)  
  Initializes the target folder as a Barretta project by creating the following structure:

  - `/barretta-core`
  - `/code_modules`
  - `/excel_file`
  - `.gitignore`
  - `barretta.code-workspace`
  - `barretta-launcher.json`

- **Barretta: Push** (Alt + Ctrl + I)  
  Imports the VBA module files from the `code_modules` folder into the Excel file in the `excel_file` folder.

- **Barretta: Pull** (Alt + Ctrl + E)  
  Exports the VBA modules from the Excel file in the `excel_file` folder to the `code_modules` folder.

- **Barretta: Open** (Alt + Ctrl + O)  
  Opens the Excel file in the `excel_file` folder.

## Barretta Launcher

A **BARRETTA-LAUNCHER** view appears in the Visual Studio Code sidebar.  
From here, you can run various commands and use the Macro Runner.

The Macro Runner allows you to execute macros defined in `barretta-launcher.json` within your Excel file.

### Example

1. When you run `Barretta: Init`, check that `barretta-launcher.json` contains three sample macros.
2. Place any macro-enabled XLSM file into the **excel_file** folder and paste the following code into the Workbook module:

   ```vb
   Option Explicit

   Sub SampleMacro1()
       MsgBox "Your first macro will be executed."
   End Sub

   Sub SampleMacro2(ByVal str1 As String, ByVal str2 As String, ByVal str3 As String)
       MsgBox str1 & str2 & str3
   End Sub

   Sub SampleMacro3(ByVal num1 As Long, ByVal num2 As Long)
       MsgBox num1 + num2
   End Sub
   ```

3. Execute the macro using the Run button in the Macro Runner.

### About Developing an LSP for VBA

If you are able to develop an LSP for VBA to be used with VSCode, your contribution would be greatly appreciated.
