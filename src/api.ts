import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import * as gen from "./generator";
import { setRootPath } from "./lib_vscode_api";
import {
  preCheckCallMacro,
  preCheckInit,
  preCheckOpen,
  preCheckPull,
  preCheckPush,
} from "./api_precheck";

const { spawn } = require("child_process");

type Initialize = {
  (): void;
};

type OpenBook = {
  (): void;
};

type CallMacro = {
  (
    callMethod: string,
    methodParams?: (string | number | boolean)[] | undefined
  ): void;
};

type PushExcel = {
  (): void;
};

type PullExcel = {
  (): void;
};

type RunPS1 = {
  (ps1Params: Ps1Params): Promise<boolean>;
};

type Ps1Params = {
  execType: string;
  ps1FilePath: string;
};

export const initialize: Initialize = async () => {
  console.log(`Barretta: Start initialize.`);

  const rootPath: string | undefined = await setRootPath();

  if (rootPath === undefined) {
    console.log(`Barretta: Exit initialize.`);
    return;
  }

  if (!preCheckInit(rootPath)) {
    console.log(`Barretta: Failed preCheckInit.`);
    return;
  }

  try {
    fs.appendFileSync(
      path.join(rootPath, ".gitignore"),
      gen.generateGitignore()
    );
    console.log("Barretta: CreateFile : .gitignore");

    fs.appendFileSync(
      path.join(rootPath, "barretta.code-workspace"),
      gen.generateWorkspace()
    );
    console.log("Barretta: CreateFile : barretta.code-workspace");

    fs.mkdirSync(path.join(rootPath, "excel_file"));
    console.log("Barretta: CreateFolder : excel_file");

    fs.mkdirSync(path.join(rootPath, "code_modules"));
    console.log("Barretta: CreateFolder : code_modules");

    fs.mkdirSync(path.join(rootPath, "barretta-core"));
    console.log("Barretta: CreateFolder : barretta-core");

    fs.mkdirSync(path.join(rootPath, "barretta-core/scripts"));
    console.log("Barretta: CreateFolder : scripts");

    fs.mkdirSync(path.join(rootPath, "barretta-core/dist"));
    console.log("Barretta: CreateFolder : dist");

    fs.mkdirSync(path.join(rootPath, "barretta-core/types"));
    console.log("Barretta: CreateFolder : types");

    fs.appendFileSync(
      path.join(rootPath, "barretta-launcher.json"),
      gen.generateBarrettaLauncher()
    );
    console.log("Barretta: CreateFile : barretta-launcher.json");

    vscode.window.showInformationMessage(
      "Barretta: Folder initialization completed."
    );
    console.log(`Barretta: Complete initialize.`);
  } catch {
    vscode.window.showErrorMessage("Barretta: Folder initialization failed.");
    console.log(`Barretta: Failed initialize.`);
  }
};

export const pushExcel: PushExcel = async () => {
  console.log(`Barretta: Start pushExcel.`);

  const rootPath: string | undefined = await setRootPath();

  if (rootPath === undefined) {
    console.log(`Barretta: Exit pushExcel.`);
    return;
  }

  if (!preCheckPush(rootPath)) {
    console.log(`Barretta: Failed preCheckPush.`);
    return;
  }

  const fileList: string[] = fs.readdirSync(path.join(rootPath, "excel_file"));
  const excelFileList: string[] = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  const fileName: string = excelFileList[0];

  // Read the VBA password from the "barretta" configuration.
  const barrettaConfig = vscode.workspace.getConfiguration("barretta");
  const vbaPassword: string = barrettaConfig.get("vbaPassword") || "";

  vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: "Barretta: Push" },
    async (progress) => {
      progress.report({ message: "Working...." });

      const ps1FilePath = path.join(
        rootPath,
        "barretta-core/scripts/push_modules.ps1"
      );

      try {
        // Recreate or clean up the [dist] folder.
        const distPath: string = path.join(rootPath, "barretta-core/dist");
        if (!fs.existsSync(distPath)) {
          fs.mkdirSync(distPath);
          console.log(`Barretta: Recreate [dist] folder.`);
        } else {
          fs.readdirSync(distPath).map((file) => {
            fs.unlinkSync(path.join(distPath, file));
            console.log(`Barretta: Delete ${file}.`);
          });
        }

        // Copy files from code_modules to dist.
        const codeModulesPath: string = path.join(rootPath, "code_modules");
        fs.readdirSync(codeModulesPath).map((file) => {
          fs.copyFileSync(
            path.join(codeModulesPath, file),
            path.join(distPath, file)
          );
          console.log(`Barretta: File Copied to dist. : ${file}`);
        });

        // Generate push_modules.ps1 with the new parameter.
        const configPush: vscode.WorkspaceConfiguration =
          vscode.workspace.getConfiguration("push");
        const pushIgnoreDocument: boolean =
          configPush.get("ignoreDocuments") ?? false;

        const genParams = {
          rootPath,
          fileName,
          pushIgnoreDocument,
          vbaPassword,
        };
        fs.appendFileSync(ps1FilePath, gen.generatePushPs1(genParams));
        console.log("Barretta: push_modules.ps1 Created.");

        // Execute the ps1 file.
        const ps1Params: Ps1Params = {
          execType: "-File",
          ps1FilePath,
        };

        if (await runPs1(ps1Params)) {
          vscode.window.showInformationMessage(
            `Barretta: Excel file imported.`
          );
          console.log(`Barretta: Complete pushExcel.`);
        } else {
          vscode.window.showErrorMessage(
            `Barretta: Code module import failed.`
          );
          console.log(`Barretta: Failed pushExcel.`);
        }
      } catch (e) {
        console.error(`Barretta: Unknown error has occurred.`);
        console.error(e);
      } finally {
        fs.unlinkSync(ps1FilePath);
        console.log(`Barretta: push_modules.ps1 deleted.`);
      }
    }
  );
};

export const pullExcel: PullExcel = async () => {
  console.log(`Barretta: Start pullExcel.`);

  const rootPath: string | undefined = await setRootPath();

  if (rootPath === undefined) {
    console.log(`Barretta: Exit pullExcel.`);
    return;
  }

  if (!preCheckPull(rootPath)) {
    console.log(`Barretta: Failed preCheckPull.`);
    return;
  }

  const fileList: string[] = fs.readdirSync(path.join(rootPath, "excel_file"));
  const excelFileList: string[] = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  const fileName: string = excelFileList[0];

  vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: "Barretta: Pull" },
    async (progress) => {
      progress.report({ message: "Working...." });

      const ps1FilePath = path.join(
        rootPath,
        "barretta-core/scripts/pull_modules.ps1"
      );

      try {
        // Generate pull_modules.ps1
        const config: vscode.WorkspaceConfiguration =
          vscode.workspace.getConfiguration("pull");
        const pullIgnoreDocument: boolean =
          config.get("ignoreDocuments") ?? false;

        const barrettaConfig = vscode.workspace.getConfiguration("barretta");
        const vbaPassword: string = barrettaConfig.get("vbaPassword") || "";

        const genParams = {
          rootPath,
          fileName,
          pullIgnoreDocument,
          vbaPassword,
        };
        fs.appendFileSync(ps1FilePath, gen.generatePullPs1(genParams));
        console.log("Barretta: pull_modules.ps1 Created.");

        // Execute the ps1 file.
        const ps1Params: Ps1Params = {
          execType: "-File",
          ps1FilePath,
        };

        if (await runPs1(ps1Params)) {
          vscode.window.showInformationMessage(
            "Barretta: Exported to code_modules folder."
          );
          console.log(`Barretta: Complete pullExcel.`);
        } else {
          vscode.window.showErrorMessage(
            "Barretta: Code module export failed."
          );
          console.log(`Barretta: Failed pullExcel.`);
        }
      } catch (e) {
        console.error(`Barretta: Unknown error has occurred.`);
        console.error(e);
      } finally {
        fs.unlinkSync(ps1FilePath);
        console.log(`Barretta: pull_modules.ps1 deleted.`);
      }
    }
  );
};

export const openBook: OpenBook = async () => {
  console.log(`Barretta: Start openBook.`);

  const rootPath: string | undefined = await setRootPath();

  if (rootPath === undefined) {
    console.log(`Barretta: Exit openBook.`);
    return;
  }

  if (!preCheckOpen(rootPath)) {
    console.log(`Barretta: Failed preCheckOpen.`);
    return;
  }

  const fileList: string[] = fs.readdirSync(path.join(rootPath, "excel_file"));
  const excelFileList: string[] = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  const fileName: string = excelFileList[0];

  vscode.window.withProgress(
    { location: vscode.ProgressLocation.Notification, title: "Barretta: Open" },
    async (progress) => {
      progress.report({ message: "Working...." });

      const ps1FilePath = path.join(
        rootPath,
        "barretta-core/scripts/open_excelbook.ps1"
      );

      try {
        // Generate open_excelbook.ps1
        const genParams = {
          rootPath,
          fileName,
        };
        fs.appendFileSync(ps1FilePath, gen.generateOpenBookPs1(genParams));
        console.log("Barretta: open_excelbook.ps1 Created.");

        // Execute the ps1 file.
        const ps1Params: Ps1Params = {
          execType: "-File",
          ps1FilePath,
        };

        if (await runPs1(ps1Params)) {
          vscode.window.showInformationMessage(
            "Barretta: Excel workbook opened."
          );
          console.log(`Barretta: Complete openBook.`);
        } else {
          vscode.window.showErrorMessage(
            "Barretta: Could not open Excel workbook."
          );
          console.log(`Barretta: Failed openBook.`);
        }
      } catch (e) {
        console.error(`Barretta: Unknown error has occurred.`);
        console.error(e);
      } finally {
        fs.unlinkSync(ps1FilePath);
        console.log(`Barretta: open_excelbook.ps1 deleted.`);
      }
    }
  );
};

export const callMacro: CallMacro = async (callMethod, methodParams?) => {
  console.log(`Barretta: Start callMacro.`);

  const rootPath: string | undefined = await setRootPath();

  if (rootPath === undefined) {
    console.log(`Barretta: Exit callMacro.`);
    return;
  }

  if (!preCheckCallMacro(rootPath)) {
    console.log(`Barretta: Failed callMacro.`);
    return;
  }

  const fileList: string[] = fs.readdirSync(path.join(rootPath, "excel_file"));
  const excelFileList: string[] = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  const fileName: string = excelFileList[0];

  // Generate run_macro.ps1
  const genParams = {
    rootPath,
    fileName,
    callMethod,
    methodParams,
  };
  fs.appendFileSync(
    path.join(rootPath, "barretta-core/scripts/run_macro.ps1"),
    gen.generateRunMacroPs1(genParams)
  );

  vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: "Barretta: LaunchMacro",
    },
    async (progress) => {
      progress.report({ message: "Working...." });
      try {
        const ps1FilePath = path.join(
          rootPath,
          "barretta-core/scripts/run_macro.ps1"
        );

        // Execute the ps1 file.
        const ps1Params: Ps1Params = {
          execType: "-File",
          ps1FilePath,
        };

        if (await runPs1(ps1Params)) {
          vscode.window.showInformationMessage(
            "Barretta: Macro execution completed."
          );
          console.log(`Barretta: run_macro.ps1 completed.`);
        } else {
          vscode.window.showErrorMessage("Barretta: Macro execution failed.");
          console.log(`Barretta: run_macro.ps1 failed.`);
        }
      } catch (e) {
        console.error(`Barretta: Unknown error has occurred.`);
        console.error(e);
      } finally {
        fs.unlinkSync(
          path.join(rootPath, "barretta-core/scripts/run_macro.ps1")
        );
        console.log(`Barretta: run_macro.ps1 deleted.`);
      }
    }
  );
};

const runPs1: RunPS1 = async (ps1Params): Promise<boolean> => {
  console.log(`runPs1 : start`);

  const child = spawn("powershell.exe", [
    "-ExecutionPolicy",
    "Bypass",
    ps1Params.execType,
    ps1Params.ps1FilePath,
  ]);

  let info = "";
  for await (const chunk of child.stdout) {
    console.log(`Powershell Info: ${chunk}`);
    info += chunk;
  }

  let error = "";
  for await (const chunk of child.stderr) {
    console.log(`Powershell Errors: ${chunk}`);
    error += chunk;
  }

  const exitCode = await new Promise((resolve) => {
    child.on("close", resolve);
  });

  if (exitCode) {
    const result = `SubProcess Error Exit ${exitCode}, ${error}`;
    console.log(`runPs1 : error-end ${result}`);
    throw new Error(result);
  }

  return error === "" ? true : false;
};

export const pushSingleExcel = async (moduleName: string) => {
  console.log(`Barretta: Start pushSingleExcel.`);

  const rootPath: string | undefined = await setRootPath();
  if (rootPath === undefined) {
    console.log(`Barretta: Exit pushSingleExcel.`);
    return;
  }

  if (!preCheckPush(rootPath)) {
    console.log(`Barretta: Failed preCheckPush.`);
    return;
  }

  const fileList: string[] = fs.readdirSync(path.join(rootPath, "excel_file"));
  const excelFileList: string[] = fileList.filter((fileName) =>
    fileName.match(/^(?!~\$).*\.(xls$|xlsx$|xlsm$|xlsb$|xlam$)/g)
  );
  const excelFileName: string = excelFileList[0];

  // Module file path
  const codeModulesPath = path.join(rootPath, "code_modules");
  const moduleFilePath = path.join(codeModulesPath, moduleName);
  const moduleFileName = path.basename(moduleFilePath);

  // dist folder path
  const distPath: string = path.join(rootPath, "barretta-core/dist");
  const ps1FilePath = path.join(
    rootPath,
    "barretta-core/scripts/push_modules.ps1"
  );

  // Read the VBA password from the "barretta" configuration.
  const barrettaConfig = vscode.workspace.getConfiguration("barretta");
  const vbaPassword: string = barrettaConfig.get("vbaPassword") || "";

  vscode.window.withProgress(
    {
      location: vscode.ProgressLocation.Notification,
      title: "Barretta: Push (Single Module)",
    },
    async (progress) => {
      progress.report({ message: "Working on single module push..." });

      try {
        // Recreate or clean up the [dist] folder.
        if (!fs.existsSync(distPath)) {
          fs.mkdirSync(distPath);
          console.log(`Barretta: Recreate [dist] folder.`);
        } else {
          fs.readdirSync(distPath).map((file) => {
            fs.unlinkSync(path.join(distPath, file));
            console.log(`Barretta: Delete ${file}.`);
          });
        }

        // Copy module file to dist.
        fs.copyFileSync(moduleFilePath, path.join(distPath, moduleFileName));
        console.log(
          `Barretta: Module File Copied to dist. : ${moduleFileName}`
        );

        // If the module is a form (.frm), also copy the corresponding .frx file
        if (path.extname(moduleFileName).toLowerCase() === ".frm") {
          const baseName = path.basename(moduleFileName, ".frm");
          const frxFileName = baseName + ".frx";
          const srcFrxFilePath = path.join(codeModulesPath, frxFileName);
          if (fs.existsSync(srcFrxFilePath)) {
            const destFrxFilePath = path.join(distPath, frxFileName);
            fs.copyFileSync(srcFrxFilePath, destFrxFilePath);
            console.log(
              `Barretta: Associated FRX file copied to dist: ${frxFileName}`
            );
          } else {
            console.warn(
              `Barretta: Expected FRX file not found: ${frxFileName}`
            );
          }
        }

        // Generate push_modules.ps1 with the new parameters
        const configPush: vscode.WorkspaceConfiguration =
          vscode.workspace.getConfiguration("push");
        const pushIgnoreDocument: boolean =
          configPush.get("ignoreDocuments") ?? false;
        const genParams = {
          rootPath,
          fileName: excelFileName, // excel file name remains the same
          pushIgnoreDocument,
          vbaPassword,
        };
        fs.appendFileSync(ps1FilePath, gen.generatePushPs1(genParams));
        console.log("Barretta: push_modules.ps1 Created.");

        // Execute the ps1 file
        const ps1Params: Ps1Params = {
          execType: "-File",
          ps1FilePath,
        };

        if (await runPs1(ps1Params)) {
          vscode.window.showInformationMessage(
            `Barretta: Excel file imported (single module).`
          );
          console.log(`Barretta: Complete pushSingleExcel.`);
        } else {
          vscode.window.showErrorMessage(
            `Barretta: Code module import failed (single module).`
          );
          console.log(`Barretta: Failed pushSingleExcel.`);
        }
      } catch (e) {
        console.error(
          `Barretta: Unknown error has occurred in pushSingleExcel.`
        );
        console.error(e);
      } finally {
        if (fs.existsSync(ps1FilePath)) {
          fs.unlinkSync(ps1FilePath);
          console.log(`Barretta: push_modules.ps1 deleted.`);
        }
      }
    }
  );
};
