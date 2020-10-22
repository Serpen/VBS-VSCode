import { commands, ExtensionContext, Uri, workspace } from "vscode";
import * as cmds from "./commands";
import hoverProvider from "./hover";
import completionProvider from "./completion";
import symbolsProvider from "./symbols";
import signatureProvider from "./signature";
import definitionProvider from "./definition";
import * as fs from "fs";
import * as pathns from "path";

class IncludeFile {
  constructor(path: string) {
    if (!pathns.isAbsolute(path))
      path = pathns.join(workspace.workspaceFolders[0].uri.fsPath, path.substr(1));

    this.Uri = Uri.file(path);

    if (fs.existsSync(path) && fs.statSync(path).isFile())
      this.Content = fs.readFileSync(path).toString();
  }
  Content = "";
  Uri: Uri;
}

export const Includes = new Map<string, IncludeFile>();

function reloadImportDocuments() {
  const SourceImportFiles = workspace.getConfiguration("vbs").get<string[]>("includes");
  for (const key of Includes.keys()) {
    if (key.startsWith("Import"))
      Includes.delete(key);
  }
  SourceImportFiles?.forEach((file, index) => {
    Includes.set("Import" + (index + 1), new IncludeFile(file));
  }
  );
}

export function activate(context: ExtensionContext): void {
  Includes.set("Global", new IncludeFile(context.asAbsolutePath("./GlobalDefs.vbs")));
  Includes.set("ObjectDefs", new IncludeFile(context.asAbsolutePath("./ObjectDefs.vbs")));

  workspace.onDidChangeConfiguration(reloadImportDocuments);
  reloadImportDocuments();

  context.subscriptions.push(
    hoverProvider,
    completionProvider,
    symbolsProvider,
    signatureProvider,
    definitionProvider);

  // Run Script Command
  commands.registerCommand("vbs.runScript", () => {
    cmds.runScript();
  });

  // Kill running script command
  commands.registerCommand("vbs.killScript", () => {
    cmds.killScript();
  });
}


