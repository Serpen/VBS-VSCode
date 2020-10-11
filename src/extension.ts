import { commands, ExtensionContext, workspace } from 'vscode';
import * as cmds from './commands';
import hoverProvider from './hover';
import completionProvider from './completion';
import symbolsProvider from './symbols';
import signatureProvider from './signature';
import definitionProvider from './definition';
import fs from 'fs';

export let GlobalSourceImport = '';
export let GlobalSourceImportFile = '';
export let ObjectSourceImport = '';
export let ObjectSourceImportFile = '';
export let SourceImports: string[] = [];
export let SourceImportFiles: string[] = [];

function reloadImportDocuments() {
  SourceImports = [];
  SourceImportFiles = workspace.getConfiguration("vbs").get<string[]>("includes");
  SourceImportFiles?.forEach((SourceImportFile, index) => {
    if (SourceImportFile.startsWith(".\\") && workspace.workspaceFolders) {
      SourceImportFile = workspace.workspaceFolders[0].uri.fsPath + SourceImportFile.substr(1);
      SourceImportFiles[index] = SourceImportFile;
    }
    if (SourceImportFile !== '' && fs.statSync(SourceImportFile)) {
      const ExtraDocumentText = fs.readFileSync(SourceImportFile).toString();

      SourceImports.push(ExtraDocumentText);
    }
  });
}

export function activate(context: ExtensionContext) : void {
  const providers = [
    hoverProvider,
    completionProvider,
    symbolsProvider,
    signatureProvider,
    definitionProvider,
  ];

  GlobalSourceImportFile = context.asAbsolutePath("./GlobalDefs.vbs");
  GlobalSourceImport = fs.readFileSync(GlobalSourceImportFile).toString();

  ObjectSourceImportFile = context.asAbsolutePath("./ObjectDefs.vbs");
  ObjectSourceImport = fs.readFileSync(ObjectSourceImportFile).toString();

  workspace.onDidChangeConfiguration(reloadImportDocuments);
  reloadImportDocuments();

  context.subscriptions.push(...providers);

  // Run Script Command
  commands.registerCommand('vbs.runScript', () => {
    cmds.runScript();
  });

  // Kill running script command
  commands.registerCommand('vbs.killScript', () => {
    cmds.killScript();
  });
}


