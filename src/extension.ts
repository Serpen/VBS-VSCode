import { commands, ExtensionContext, workspace } from "vscode";
import hoverProvider from "./hover";
import completionProvider from "./completion";
import symbolsProvider from "./symbols";
import signatureProvider from "./signature";
import definitionProvider from "./definition";
import colorProvider from "./colorprovider";
import launchProvider from "./Launcher";
import formatter from "./formatter";
import * as cmds from "./commands";
import { IncludeFile, Includes } from "./Includes";
import { basename } from "path";

function reloadImportDocuments() {
  const SourceImportFiles = workspace.getConfiguration("vbs").get<string[]>("includes");
  for (const key of Includes.keys()) {
    if (key.startsWith("Import"))
      Includes.delete(key);
  }
  SourceImportFiles?.forEach((file) => {
    Includes.set(`Import ${basename(file)}`, new IncludeFile(file));
  });
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
    definitionProvider,
    colorProvider,
    formatter,
    launchProvider.launchConfigProvider,
    launchProvider.inlineDebugAdapterFactory
  );

  // Run Script Command
  commands.registerCommand("vbs.runScript", () => {
    cmds.runScript();
  });

  // Kill running script command
  commands.registerCommand("vbs.killScript", () => {
    cmds.killScript();
  });
}
