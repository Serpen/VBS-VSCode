import { commands, ExtensionContext  } from 'vscode';
// import langConfig from '../syntaxes/language-configuration.json';
import * as AutoItCommands from './ai_commands';
import * as hoverFeature from './ai_hover';
import * as completionFeature from './ai_completion';
import * as symbolsFeature from './ai_symbols';
import * as signaturesFeature from './ai_signature';
import * as workspaceSymbolsFeature from './ai_workspaceSymbols';
import * as goToDefinitionFeature from './ai_definition';

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

  // context.subscriptions.push(languages.setLanguageConfiguration('autoit', langConfig));

  // Run Script Command
  commands.registerCommand('extension.runScript', () => {
    AutoItCommands.runScript();
  });

  // Launch Debug-Console
  commands.registerCommand('extension.debugConsole', () => {
    AutoItCommands.debugConsole();
  });


  // Update console parameters
  commands.registerCommand('extension.changeParams', () => {
    AutoItCommands.changeConsoleParams();
  });

  // Kill running script command
  commands.registerCommand('extension.killScript', () => {
    AutoItCommands.killScript();
  });

  // eslint-disable-next-line no-console
  console.log('AutoIt is now active!');
}

// this method is called when your extension is deactivated
export function deactivate() {}

