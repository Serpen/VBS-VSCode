import { Diagnostic, DiagnosticSeverity, languages, Range, window, workspace } from 'vscode';
import { ChildProcessWithoutNullStreams, spawn } from 'child_process';
import path from 'path';
import localize from './localize';

const configuration = workspace.getConfiguration('vbs');

const vbsOut = window.createOutputChannel('VBScript');

let runner: ChildProcessWithoutNullStreams;

const scriptInterpreter: string = configuration.get<string>("interpreter");

const dc = languages.createDiagnosticCollection("vbs");

function procRunner(cmdPath: string, args: string[]) {
  vbsOut.clear();
  vbsOut.show(true);

  dc.clear();

  if (!window.activeTextEditor)
    return;

  const workDir = path.dirname(window.activeTextEditor.document.fileName);

  runner = spawn(cmdPath, args, {
    cwd: workDir,
  });

  runner.stdout.on('data', data => {
    const output = data.toString();
    vbsOut.append(output);
  });

  runner.stderr.on('data', data => {
    const output = data.toString();
    let match = (/.*\((\d+), (\d+)\) (.*)/.exec(output));
    if (match) {
      const line = Number.parseInt(match[1]) - 1;
      const char = Number.parseInt(match[2]) - 1;
      const diag = new Diagnostic(new Range(line, char, line, char), match[3], DiagnosticSeverity.Error);
      dc.set(window.activeTextEditor.document.uri, [diag]);
    }
    vbsOut.append(output);
  });

  runner.on('exit', code => {
    vbsOut.appendLine(`Process exited with code ${code}`);
  });
}

export function runScript(): void {
  if (!window.activeTextEditor)
    return;
  const thisDoc = window.activeTextEditor.document; // Get the object of the text editor
  // Save the file
  thisDoc.save().then(() => {
    window.setStatusBarMessage(localize("vbs.runningscript"), 1500);

    procRunner(scriptInterpreter, [thisDoc.fileName]);
  });
}

export function killScript(): void {
  // runner.stdin.pause();
  runner?.kill();
}
