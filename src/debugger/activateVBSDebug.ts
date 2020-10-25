import * as vscode from "vscode";
import { MockDebugSession } from "./vbsDebug";

class ConfigurationProvider implements vscode.DebugConfigurationProvider {
  resolveDebugConfiguration(_folder: vscode.WorkspaceFolder | undefined, config: vscode.DebugConfiguration)
    : vscode.ProviderResult<vscode.DebugConfiguration> {

    const editor = vscode.window.activeTextEditor;
    if (editor && editor.document.languageId === "vbs") {
      config.type = "vbs";
      config.name = "Launch";
      config.request = "launch";
      config.program = "${file}";
      config.stopOnEntry = true;
    }
    
    if (!config.program) {
      return vscode.window.showInformationMessage("Cannot find a program to debug").then(() => {
        return undefined;	// abort launch
      });
    }

    return config;
  }
}

class InlineDebugAdapterFactory implements vscode.DebugAdapterDescriptorFactory {

  createDebugAdapterDescriptor(): vscode.ProviderResult<vscode.DebugAdapterDescriptor> {
    return new vscode.DebugAdapterInlineImplementation(new MockDebugSession());
  }
}

export function activateMockDebug(context: vscode.ExtensionContext) : void {
  // register a configuration provider for 'mock' debug type
  const cfgprovider = new ConfigurationProvider();
  context.subscriptions.push(vscode.debug.registerDebugConfigurationProvider("vbs", cfgprovider));

  const dbafactory = new InlineDebugAdapterFactory();
  
  context.subscriptions.push(vscode.debug.registerDebugAdapterDescriptorFactory("vbs", dbafactory));
}
