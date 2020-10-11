import { window, workspace } from 'vscode';
import { ChildProcessWithoutNullStreams, spawn } from 'child_process';
import path from 'path';

const configuration = workspace.getConfiguration('vbs');

const vbsOut = window.createOutputChannel('VBScript');

let runner: ChildProcessWithoutNullStreams;

const scriptInterpreter: string = configuration.get<string>("interpreter");

function procRunner(cmdPath: string, args: string[]) {
  vbsOut.clear();
  vbsOut.show(true);

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
    vbsOut.append(output);
  });

  runner.on('exit', code => {
    vbsOut.appendLine(`Process exited with code ${code}`);
  });
}

export function runScript() : void {
  if (!window.activeTextEditor)
    return;
  const thisDoc = window.activeTextEditor.document; // Get the object of the text editor
  // Save the file
  thisDoc.save().then(() => {
    window.setStatusBarMessage('Running the script...', 1500);

    procRunner(scriptInterpreter, [thisDoc.fileName]);
  });
}

export function killScript() : void {
  // runner.stdin.pause();
  runner?.kill();
}
