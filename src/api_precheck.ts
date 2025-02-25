import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";

type PreCheckInit = {
  (rootPath: string): boolean;
};

type PreCheckPush = {
  (rootPath: string): boolean;
};

type PreCheckPull = {
  (rootPath: string): boolean;
};

type PreCheckOpen = {
  (rootPath: string): boolean;
};

type PreCheckCallMacro = {
  (rootPath: string): boolean;
};

export const preCheckInit: PreCheckInit = (rootPath) => {
  // Check if [barretta-core] folder exists.
  const barretta_core = path.join(rootPath, "barretta-core");
  if (fs.existsSync(barretta_core)) {
    vscode.window.showErrorMessage(
      `Barretta: The target folder is already initialized.\n${rootPath}`
    );
    console.log(`Barretta: The target folder has already been initialized.`);
    return false;
  }

  return true;
};

export const preCheckPush: PreCheckPush = (rootPath) => {
  console.log(`Barretta: Start preCheckPush.`);

  // Check if [barretta-core] folder exists.
  const barretta_core = path.join(rootPath, "barretta-core");
  if (!fs.existsSync(barretta_core)) {
    vscode.window.showErrorMessage(
      `Barretta: The target folder does not contain a [barretta-core] folder.\nPlease run the [Barretta: Init] command to create the Barretta project.\n${barretta_core}`
    );
    console.log(`Barretta: [barretta-core] folder does not exist.`);
    return false;
  }

  // Check if [excel_file] folder exists.
  const excelFolderPath: string = path.join(rootPath, "excel_file");
  if (!fs.existsSync(excelFolderPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [excel_file] folder does not exist.\n${excelFolderPath}`
    );
    console.log(`Barretta: [excel_file] folder does not exist.`);
    return false;
  }

  // Check that a single Excel file exists.
  const fileList = fs.readdirSync(excelFolderPath);
  const excelFileList = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  if (excelFileList.length === 0) {
    vscode.window.showErrorMessage(`Barretta: No Excel file is found.`);
    console.log(`Barretta: Excel file does not exist.`);
    return false;
  } else if (excelFileList.length > 1) {
    vscode.window.showErrorMessage(`Barretta: Multiple Excel files found.`);
    console.log(`Barretta: Multiple Excel files exist.`);
    return false;
  }

  // Check if [code_modules] folder exists.
  const codeModulesPath = path.join(rootPath, "code_modules");
  if (!fs.existsSync(codeModulesPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [code_modules] folder does not exist.\n${codeModulesPath}`
    );
    console.log(`Barretta: [code_modules] folder does not exist.`);
    return false;
  }

  // Check if [scripts] folder exists.
  const scriptsPath = path.join(rootPath, "barretta-core/scripts");
  if (!fs.existsSync(scriptsPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [scripts] folder does not exist.\n${scriptsPath}`
    );
    console.log(`Barretta: [scripts] folder does not exist.`);
    return false;
  }

  console.log(`Barretta: Complete preCheckPush.`);
  return true;
};

export const preCheckPull: PreCheckPull = (rootPath) => {
  console.log(`Barretta: Start preCheckPull.`);

  const barretta_core = path.join(rootPath, "barretta-core");
  if (!fs.existsSync(barretta_core)) {
    vscode.window.showErrorMessage(
      `Barretta: The target folder does not contain a [barretta-core] folder.\nPlease run the [Barretta: Init] command to create the Barretta project.\n${barretta_core}`
    );
    console.log(`Barretta: [barretta-core] folder does not exist.`);
    return false;
  }

  const excelFolderPath: string = path.join(rootPath, "excel_file");
  if (!fs.existsSync(excelFolderPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [excel_file] folder does not exist.\n${excelFolderPath}`
    );
    console.log(`Barretta: [excel_file] folder does not exist.`);
    return false;
  }

  const fileList = fs.readdirSync(excelFolderPath);
  const excelFileList = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  if (excelFileList.length === 0) {
    vscode.window.showErrorMessage(`Barretta: No Excel file is found.`);
    console.log(`Barretta: Excel file does not exist.`);
    return false;
  } else if (excelFileList.length > 1) {
    vscode.window.showErrorMessage(`Barretta: Multiple Excel files found.`);
    console.log(`Barretta: Multiple Excel files exist.`);
    return false;
  }

  if (!fs.existsSync(codeModulesPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [code_modules] folder does not exist.\n${codeModulesPath}`
    );
    console.log(`Barretta: [code_modules] folder does not exist.`);
    return false;
  }

  const scriptsPath = path.join(rootPath, "barretta-core/scripts");
  if (!fs.existsSync(scriptsPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [scripts] folder does not exist.\n${scriptsPath}`
    );
    console.log(`Barretta: [scripts] folder does not exist.`);
    return false;
  }

  console.log(`Barretta: Complete preCheckPull.`);
  return true;
};

export const preCheckOpen: PreCheckOpen = (rootPath) => {
  console.log(`Barretta: Start preCheckOpen.`);

  const barretta_core = path.join(rootPath, "barretta-core");
  if (!fs.existsSync(barretta_core)) {
    vscode.window.showErrorMessage(
      `Barretta: The target folder does not contain a [barretta-core] folder.\nPlease run the [Barretta: Init] command to create the Barretta project.\n${barretta_core}`
    );
    console.log(`Barretta: [barretta-core] folder does not exist.`);
    return false;
  }

  const excelFolderPath: string = path.join(rootPath, "excel_file");
  if (!fs.existsSync(excelFolderPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [excel_file] folder does not exist.\n${excelFolderPath}`
    );
    console.log(`Barretta: [excel_file] folder does not exist.`);
    return false;
  }

  const fileList = fs.readdirSync(excelFolderPath);
  const excelFileList = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  if (excelFileList.length === 0) {
    vscode.window.showErrorMessage(`Barretta: No Excel file is found.`);
    console.log(`Barretta: Excel file does not exist.`);
    return false;
  } else if (excelFileList.length > 1) {
    vscode.window.showErrorMessage(`Barretta: Multiple Excel files found.`);
    console.log(`Barretta: Multiple Excel files exist.`);
    return false;
  }

  const scriptsPath = path.join(rootPath, "barretta-core/scripts");
  if (!fs.existsSync(scriptsPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [scripts] folder does not exist.\n${scriptsPath}`
    );
    console.log(`Barretta: [scripts] folder does not exist.`);
    return false;
  }

  console.log(`Barretta: Complete preCheckOpen.`);
  return true;
};

export const preCheckCallMacro: PreCheckCallMacro = (rootPath) => {
  console.log(`Barretta: Start preCheckCallMacro.`);

  const barretta_core = path.join(rootPath, "barretta-core");
  if (!fs.existsSync(barretta_core)) {
    vscode.window.showErrorMessage(
      `Barretta: The target folder does not contain a [barretta-core] folder.\nPlease run the [Barretta: Init] command to create the Barretta project.\n${barretta_core}`
    );
    console.log(`Barretta: [barretta-core] folder does not exist.`);
    return false;
  }

  const excelFolderPath: string = path.join(rootPath, "excel_file");
  if (!fs.existsSync(excelFolderPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [excel_file] folder does not exist.\n${excelFolderPath}`
    );
    console.log(`Barretta: [excel_file] folder does not exist.`);
    return false;
  }

  const fileList = fs.readdirSync(excelFolderPath);
  const excelFileList = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  if (excelFileList.length === 0) {
    vscode.window.showErrorMessage(`Barretta: No Excel file is found.`);
    console.log(`Barretta: Excel file does not exist.`);
    return false;
  } else if (excelFileList.length > 1) {
    vscode.window.showErrorMessage(`Barretta: Multiple Excel files found.`);
    console.log(`Barretta: Multiple Excel files exist.`);
    return false;
  }

  if (!fs.existsSync(codeModulesPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [code_modules] folder does not exist.\n${codeModulesPath}`
    );
    console.log(`Barretta: [code_modules] folder does not exist.`);
    return false;
  }

  const scriptsPath = path.join(rootPath, "barretta-core/scripts");
  if (!fs.existsSync(scriptsPath)) {
    vscode.window.showErrorMessage(
      `Barretta: The [scripts] folder does not exist.\n${scriptsPath}`
    );
    console.log(`Barretta: [scripts] folder does not exist.`);
    return false;
  }

  console.log(`Barretta: Complete preCheckCallMacro.`);
  return true;
};
