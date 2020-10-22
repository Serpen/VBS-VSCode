import { Diagnostic, DiagnosticSeverity, Disposable, languages, Range, window, workspace } from "vscode";
import * as childProcess from "child_process";
import path from "path";
import localize from "./localize";
import * as fs from "fs";

const configuration = workspace.getConfiguration("vbs");

const vbsOut = window.createOutputChannel("VBScript");

let runner: childProcess.ChildProcessWithoutNullStreams;

const scriptInterpreter: string = configuration.get<string>("interpreter");

const diagCollection = languages.createDiagnosticCollection("vbs");

let statbar: Disposable;

export function runScript(): void {
  if (!window.activeTextEditor)
    return;

  try {
    fs.accessSync(scriptInterpreter, fs.constants.X_OK);
  } catch {
    window.showErrorMessage(localize("vbs.msg.interpreterRunError") + " " + scriptInterpreter);
  }

  0x12

  diagCollection.clear();

  const doc = window.activeTextEditor.document;
  doc.save().then(() => {
    vbsOut.clear();
    vbsOut.show(true);

    const workDir = path.dirname(doc.fileName);

    if (statbar)
      statbar.dispose();

    statbar = window.setStatusBarMessage(localize("vbs.msg.runningscript"));

    runner = childProcess.spawn(scriptInterpreter, [doc.fileName], {
      cwd: workDir,
    });

    runner.stdout.on("data", data => {
      const output = data.toString();
      vbsOut.append(output);
    });

    runner.stderr.on("data", data => {
      const output = data.toString();
      const match = (/.*\((\d+), (\d+)\) (.*)/.exec(output));
      if (match) {
        const line = Number.parseInt(match[1]) - 1;
        const char = Number.parseInt(match[2]) - 1;
        const diag = new Diagnostic(new Range(line, char, line, char), match[3], DiagnosticSeverity.Error);
        diagCollection.set(doc.uri, [diag]);
      }
      vbsOut.append(output);
    });

    runner.on("exit", code => {
      vbsOut.appendLine(`Process exited with code ${code}`);
      statbar.dispose();
    });
  }, () => {
    window.showErrorMessage("Document can' be saved");
    return;
  });
}

export function killScript(): void {
  // runner.stdin.pause();
  runner?.kill();
  statbar?.dispose();
}
