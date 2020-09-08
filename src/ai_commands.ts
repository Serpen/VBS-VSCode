import { window, Position, workspace } from 'vscode';
import {spawn } from 'child_process';
import path from 'path';

const configuration = workspace.getConfiguration('vbs');

const vbsOut = window.createOutputChannel('VBScript');

let runner;
const cscript = "C:\\WINDOWS\\system32\\cscript.exe";

function procRunner(cmdPath : string, args) {
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

const runScript = () => {
  const thisDoc = window.activeTextEditor.document; // Get the object of the text editor
  const thisFile = thisDoc.fileName; // Get the current file name

  // Save the file
  thisDoc.save();

  window.setStatusBarMessage('Running the script...', 1500);

  procRunner(cscript, [thisFile]);
};

function getDebugText() {
  const editor = window.activeTextEditor;
  const thisDoc = editor.document;
  const varToDebug = thisDoc.getText(thisDoc.getWordRangeAtPosition(editor.selection.active));

  // Make sure that a variable or macro is selected
  if (varToDebug.charAt(0) === '$' || varToDebug.charAt(0) === '@') {
    const nextLine = editor.selection.active.line + 1;
    const newPosition = new Position(nextLine, 0);

    return {
      text: varToDebug,
      position: newPosition,
    };
  }
  window.showErrorMessage(
    `"${varToDebug}" is not a variable or macro, debug line can't be generated`,
  );
  return {};
}

function getIndent() {
  const editor = window.activeTextEditor;
  const doc = editor.document;

  // Grab the whole line
  const currentLine = doc.lineAt(editor.selection.active.line).text;
  // Get the indent of the current line
  const findIndent = /(\s*).+/;
  return findIndent.exec(currentLine)[1];
}

const debugConsole = () => {
  const editor = window.activeTextEditor;
  const debugText = getDebugText();
  const indent = getIndent();

  if (debugText) {
    const debugCode = `${indent};### Debug CONSOLE ↓↓↓\n${indent}ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : ${debugText.text} = ' & ${debugText.text} & @CRLF & '>Error code: ' & @error & @CRLF)\n`;

    // Insert the code for the MsgBox into the script
    editor.edit(edit => {
      edit.insert(debugText.position, debugCode);
    });
  }
};

const killScript = () => {
  runner.stdin.pause();
  runner.kill();
};

export {
  debugConsole,
  killScript,
  runScript,
};
