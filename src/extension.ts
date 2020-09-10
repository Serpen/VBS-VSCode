import { commands, ExtensionContext, DocumentSelector  } from 'vscode';
import * as Commands from './commands';
import * as hoverFeature from './hover';
import * as completionFeature from './completion';
import * as symbolsFeature from './symbols';
import * as signaturesFeature from './signature';
import * as workspaceSymbolsFeature from './workspaceSymbols';
import * as goToDefinitionFeature from './definition';

export function activate(context: ExtensionContext) {
  const features : any = [
    hoverFeature,
    completionFeature,
    symbolsFeature,
    signaturesFeature,
    workspaceSymbolsFeature,
    goToDefinitionFeature,
  ];
  context.subscriptions.push(...features);

  // context.subscriptions.push(languages.setLanguageConfiguration('vbs', langConfig));

  // Run Script Command
  commands.registerCommand('extension.runScript', () => {
    Commands.runScript();
  });

  // Launch Debug-Console
  commands.registerCommand('extension.debugConsole', () => {
    Commands.debugConsole();
  });

  // Kill running script command
  commands.registerCommand('extension.killScript', () => {
    Commands.killScript();
  });
}

// this method is called when your extension is deactivated
export function deactivate() {}

