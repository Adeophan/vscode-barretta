import * as vscode from "vscode";

type QuickPick = {
  (listItems: string[], title: string): Promise<string | undefined>;
};

type SetRootPath = {
  (): Promise<string | undefined>;
};

export const quickPick: QuickPick = async (listItems, title) => {
  return await vscode.window.showQuickPick(listItems, { title });
};

export const setRootPath: SetRootPath = async () => {
  let rootPath;

  if (!vscode.window.activeTextEditor) {
    if (vscode.workspace.workspaceFolders === undefined) {
      vscode.window.showErrorMessage(
        `Barretta: Please run this command with a folder open.`
      );
      console.log(
        `Barretta: The target could not be identified because no folder is open.`
      );
      return rootPath;
    }

    const folders: readonly vscode.WorkspaceFolder[] =
      vscode.workspace.workspaceFolders;
    const listItems: string[] = folders.map((folder) => folder.uri.fsPath);
    rootPath = await quickPick(listItems, `Please select the target folder.`);
    if (rootPath === undefined)
      console.log(`Barretta: The root folder selection has been canceled.`);

    return rootPath;
  } else {
    const activeEditorPath: vscode.Uri =
      vscode.window.activeTextEditor.document.uri;
    if (vscode.workspace.getWorkspaceFolder(activeEditorPath) !== undefined) {
      const rootWsFolder: vscode.WorkspaceFolder | undefined =
        vscode.workspace.getWorkspaceFolder(activeEditorPath);

      rootPath = rootWsFolder && rootWsFolder.uri.path.replace(/^\//, "");
      console.log(
        `Barretta: ${rootPath} was selected from the current active editor.`
      );
      return rootPath;
    } else {
      vscode.window.showErrorMessage(
        `Barretta: Could not determine the target folder.`
      );
      console.log(`Barretta: Failed to locate root folder.`);
      return rootPath;
    }
  }
};
