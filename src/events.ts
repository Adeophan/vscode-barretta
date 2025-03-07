import * as vscode from "vscode";
import * as path from "path";
import { pushSingleExcel } from "./api";
import * as fs from "fs";
import { setRootPath } from "./lib_vscode_api";

type OnSavedCodeFile = {
  (document: vscode.TextDocument): void;
};

export const onSavedCodeFile: OnSavedCodeFile = (document) => {
  // If [barretta.hotReload] is false, process ends.
  const config: vscode.WorkspaceConfiguration =
    vscode.workspace.getConfiguration("barretta");
  const optHotReload: boolean = config.get("hotReload") ?? false;
  if (!optHotReload) return;

  // If the saved file is not a file, process ends.
  if (document.uri.scheme !== "file") return;

  // Executes Barretta: Push Single Module if the saved file is a file in 'code_modules'.
  if (vscode.workspace.getWorkspaceFolder(document.uri) !== undefined) {
    const filePath = document.uri.fsPath;
    if (path.basename(path.dirname(filePath)) === "code_modules") {
      // Use an async IIFE to handle the async operations
      (async () => {
        const rootPath: string | undefined = await setRootPath();
        if (!rootPath) return;

        // Get the excel file name from the excel_file folder
        const excelFolder = path.join(rootPath, "excel_file");
        if (!fs.existsSync(excelFolder)) return;
        const fileList: string[] = fs.readdirSync(excelFolder);
        const excelFiles: string[] = fileList.filter((fileName) =>
          fileName.match(/^(?!~\$).*(\.xls$|\.xlsx$|\.xlsm$|\.xlsb$|\.xlam$)/i)
        );
        if (excelFiles.length === 0) return;
        const excelFileName: string = excelFiles[0];

        // Read the VBA password from the 'barretta' configuration
        const barrettaConfig = vscode.workspace.getConfiguration("barretta");
        const vbaPassword: string = barrettaConfig.get("vbaPassword") || "";

        // Extract moduleName from filePath
        const moduleName = path.basename(filePath);

        // Call pushSingleExcel with only moduleName
        await pushSingleExcel(moduleName);
      })();
    }
  }
};
