import { commands, ExtensionContext } from 'vscode';
import * as cmds from './commands';
import hoverProvider from './hover';
import completionProvider from './completion';
import symbolsProvider from './symbols';
import signatureProvider from './signature';
import definitionProvider from './definition';

export function activate(context: ExtensionContext) {
  const providers: any = [
    hoverProvider,
    completionProvider,
    symbolsProvider,
    signatureProvider,
    definitionProvider,
  ];

  context.subscriptions.push(...providers);

  // Run Script Command
  commands.registerCommand('extension.runScript', () => {
    cmds.runScript();
  });

  // Kill running script command
  commands.registerCommand('extension.killScript', () => {
    cmds.killScript();
  });
}

// this method is called when your extension is deactivated
export function deactivate() { }

