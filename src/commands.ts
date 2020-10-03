import { window, workspace } from 'vscode';
import { ChildProcessWithoutNullStreams, spawn } from 'child_process';
import path from 'path';

const configuration = workspace.getConfiguration('vbs');

const vbsOut = window.createOutputChannel('VBScript');

let runner: ChildProcessWithoutNullStreams;

const cscript: string = configuration.get("interpreter");

function procRunner(cmdPath: string, args: string[]) {
  vbsOut.clear();
  vbsOut.show(true);

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

export function runScript() {
  const thisDoc = window.activeTextEditor.document; // Get the object of the text editor
  // Save the file
  thisDoc.save();

  window.setStatusBarMessage('Running the script...', 1500);

  procRunner(cscript, [thisDoc.fileName]);
}

export function killScript() {
  // runner.stdin.pause();
  runner?.kill();
}
