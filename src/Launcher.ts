// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { TerminatedEvent, OutputEvent, LoggingDebugSession } from "vscode-debugadapter";
import * as vscode from "vscode";
import { DebugProtocol } from "vscode-debugprotocol";
import { dirname } from "path";
import { ChildProcessWithoutNullStreams, spawn } from "child_process";

interface ILaunchRequestArguments extends DebugProtocol.LaunchRequestArguments {
  program: string;
}

const diagCollection = vscode.languages.createDiagnosticCollection("vbs");

export class VbsDebugSession extends LoggingDebugSession {

  private _runner : ChildProcessWithoutNullStreams;

  public constructor() {
    super();
  }

  protected initializeRequest(response: DebugProtocol.InitializeResponse): void {

    response.body = response.body || {};
    response.body.supportsConfigurationDoneRequest = true;
    response.body.supportsCancelRequest = false;
    response.body.supportsTerminateRequest = true;

    this.sendResponse(response);
  }

  protected async launchRequest(response: DebugProtocol.LaunchResponse, args: ILaunchRequestArguments) : Promise<void> {
    diagCollection.clear();

    const workDir = dirname(args.program);

    this.sendResponse(response);

    const configuration = vscode.workspace.getConfiguration("vbs");
    const scriptInterpreter: string = configuration.get<string>("interpreter");

    this._runner = spawn(scriptInterpreter, [args.program], { "cwd": workDir });

    this._runner.stdout.on("data", data => {
      this.sendEvent(new OutputEvent(data.toString(), "stdout"));
    });

    this._runner.stderr.on("data", data => {
      const output = data.toString();
      const match = (/.*\((\d+), (\d+)\) (.*)/).exec(output);
      if (match) {
        const line = Number.parseInt(match[1]) - 1;
        const char = Number.parseInt(match[2]) - 1;
        const diag = new vscode.Diagnostic(new vscode.Range(line, char, line, char), match[3], vscode.DiagnosticSeverity.Error);
        diagCollection.set(vscode.Uri.file(args.program), [diag]);
      }
      this.sendEvent(new OutputEvent(output, "stderr"));
    });

    this._runner.on("exit", code => {
      this.sendEvent(new OutputEvent(`Process [${this._runner.pid}] exited with code ${code} (&H${code.toString(16).toUpperCase()})`));
      this.sendEvent(new TerminatedEvent());
    });
  }

  protected terminateRequest(response: DebugProtocol.TerminateResponse, _args: DebugProtocol.TerminateArguments, _req?: DebugProtocol.Request): void {
    this._runner?.kill();
    this.sendEvent(new TerminatedEvent());
    this.sendResponse(response);
  }
}


export class DebugConfigurationProvider implements vscode.DebugConfigurationProvider {



  public resolveDebugConfiguration(_folder: vscode.WorkspaceFolder | undefined, config: vscode.DebugConfiguration)
  : vscode.ProviderResult<vscode.DebugConfiguration> {

    const editor = vscode.window.activeTextEditor;
    if (!config || !config.request) {
      if (editor && editor.document.languageId === "vbs") {
        config.type = "vbs";
        config.name = "CScript";
        config.request = "launch";
        config.program = editor.document.uri.fsPath;
      } else
        return;
    }

    if (!config.program || !editor || editor.document.languageId !== "vbs") {
      return vscode.window.showInformationMessage("Cannot find a program to debug").then(() => {
        return undefined;	// abort launch
      });
    }

    return config;
  }
}

export class InlineDebugAdapterFactory implements vscode.DebugAdapterDescriptorFactory {
  createDebugAdapterDescriptor(): vscode.ProviderResult<vscode.DebugAdapterDescriptor> {
    return new vscode.DebugAdapterInlineImplementation(new VbsDebugSession());
  }
}

const launchConfigProvider = vscode.debug.registerDebugConfigurationProvider("vbs", new DebugConfigurationProvider());
const inlineDebugAdapterFactory = vscode.debug.registerDebugAdapterDescriptorFactory("vbs", new InlineDebugAdapterFactory());

export default { launchConfigProvider, inlineDebugAdapterFactory };

