import {
  LoggingDebugSession,
  InitializedEvent, TerminatedEvent, OutputEvent,
} from "vscode-debugadapter";
import { DebugProtocol } from "vscode-debugprotocol";
import { dirname } from "path";
import { ChildProcessWithoutNullStreams, spawn } from "child_process";
import { Diagnostic, DiagnosticSeverity, languages, Range, Uri, workspace } from "vscode";

/**
 * This interface describes the mock-debug specific launch attributes
 * (which are not part of the Debug Adapter Protocol).
 * The schema for these attributes lives in the package.json of the mock-debug extension.
 * The interface should always match this schema.
 */
interface ILaunchRequestArguments extends DebugProtocol.LaunchRequestArguments {
  /** An absolute path to the "program" to debug. */
  program: string;
}

export class MockDebugSession extends LoggingDebugSession {

private _runner : ChildProcessWithoutNullStreams;

/**
 * Creates a new debug adapter that is used for one debug session.
 * We configure the default implementation of a debug adapter here.
 */
public constructor() {
  super("vbs-debug.txt");
}

/**
* The 'initialize' request is the first request called by the frontend
* to interrogate the features the debug adapter provides.
*/
protected initializeRequest(response: DebugProtocol.InitializeResponse): void {

  // build and return the capabilities of this debug adapter:
  response.body = response.body || {};

  // the adapter implements the configurationDoneRequest.
  response.body.supportsConfigurationDoneRequest = true;

  // make VS Code to send cancelRequests
  // response.body.supportsCancelRequest = true;
  

  this.sendResponse(response);

  // since this debug adapter can accept configuration requests like 'setBreakpoint' at any time,
  // we request them early by sending an 'initializeRequest' to the frontend.
  // The frontend will end the configuration sequence by calling 'configurationDone' request.
  this.sendEvent(new InitializedEvent());
}
private diagCollection = languages.createDiagnosticCollection("vbs");

protected async launchRequest(response: DebugProtocol.LaunchResponse, args: ILaunchRequestArguments) : Promise<void> {
  this.diagCollection.clear();

  const workDir = dirname(args.program);

  this.sendResponse(response);

  const configuration = workspace.getConfiguration("vbs");
  const scriptInterpreter: string = configuration.get<string>("interpreter");

  this._runner = spawn(scriptInterpreter, [args.program], {"cwd": workDir});

  this._runner.stdout.on("data", data => {
    this.sendEvent(new OutputEvent(`${data.toString()}`));
  });

  this._runner.stderr.on("data", data => {
    const output = data.toString();
    const match = (/.*\((\d+), (\d+)\) (.*)/.exec(output));
    if (match) {
      const line = Number.parseInt(match[1]) - 1;
      const char = Number.parseInt(match[2]) - 1;
      const diag = new Diagnostic(new Range(line, char, line, char), match[3], DiagnosticSeverity.Error);
      this.diagCollection.set(Uri.file(args.program), [diag]);
    }
    this.sendEvent(new OutputEvent(`${data}`));
  });

  this._runner.on("exit", code => {
    this.sendEvent(new OutputEvent(`Process exited with code ${code}`));
    this.sendEvent(new TerminatedEvent());
  });
}

protected cancelRequest() : void {
  this._runner?.kill();
}
}

