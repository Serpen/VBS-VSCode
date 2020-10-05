import { commands, ExtensionContext, workspace } from 'vscode';
import * as cmds from './commands';
import hoverProvider from './hover';
import completionProvider from './completion';
import symbolsProvider from './symbols';
import signatureProvider from './signature';
import definitionProvider from './definition';
import fs from 'fs';

export let GlobalSourceImport: string = '';
export let GlobalSourceImportFile: string = '';
export let ObjectSourceImport: string = '';
export let ObjectSourceImportFile: string = '';
export let SourceImports : string[] = [];
export let SourceImportFiles : string[] = [];

export let ImportDocuments = [ //doesn't work
  GlobalSourceImport,
  ...SourceImports
]

function reloadImportDocuments() {
  SourceImports = [];
  SourceImportFiles = workspace.getConfiguration("vbs").get<string[]>("includes")!;
  SourceImportFiles?.forEach((SourceImportFile, index) => {
    if (SourceImportFile.startsWith(".\\") && workspace.workspaceFolders) {
      SourceImportFile = workspace.workspaceFolders[0].uri.fsPath + SourceImportFile.substr(1);
      SourceImportFiles[index] = SourceImportFile;
    }
    if (SourceImportFile != '' && fs.statSync(SourceImportFile)) {
      const ExtraDocumentText = fs.readFileSync(SourceImportFile).toString();

      SourceImports.push(ExtraDocumentText);
    }
  });
}

export function activate(context: ExtensionContext) {
  const providers: any = [
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

// this method is called when your extension is deactivated
export function deactivate() { }

