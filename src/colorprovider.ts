import * as vscode from "vscode";
import * as PATTERNS from "./patterns";

class VBSColorProvider implements vscode.DocumentColorProvider {

  public provideDocumentColors(document: vscode.TextDocument): vscode.ColorInformation[] {
    const array = new Array<vscode.ColorInformation>();
    let matches: RegExpExecArray;
    while ((matches = PATTERNS.COLOR.exec(document.getText())) !== null) {
      const pos = document.positionAt(matches.index);
      const posEnd = document.positionAt(matches.index + matches[0].length);
      const range = new vscode.Range(pos, posEnd);

      let color: vscode.Color;

      if (matches[1]) {
        switch (matches[1]) {
        case "vbBlack":
          color = new vscode.Color(0, 0, 0, 1);
          break;
        case "vbBlue":
          color = new vscode.Color(0, 0, 1, 1);
          break;
        case "vbCyan":
          color = new vscode.Color(0, 1, 1, 1);
          break;
        case "vbGreen":
          color = new vscode.Color(0, 1, 0, 1);
          break;
        case "vbMagenta":
          color = new vscode.Color(1, 0, 1, 1);
          break;
        case "vbRed":
          color = new vscode.Color(1, 0, 0, 1);
          break;
        case "vbWhite":
          color = new vscode.Color(1, 1, 1, 1);
          break;
        case "vbYellow":
          color = new vscode.Color(1, 1, 0, 1);
          break;
        }
      } else if (matches[2]) {
        let r: number, g: number, b: number;
        if (matches[3].toLowerCase().startsWith("&h"))
          r = Number.parseInt(matches[3].substr(2), 16) / 0xff;
        else 
          r = Number.parseInt(matches[3]) / 0xff;

        if (matches[4].toLowerCase().startsWith("&h"))
          g = Number.parseInt(matches[4].substr(2), 16) / 0xff;
        else 
          g = Number.parseInt(matches[4]) / 0xff;

        if (matches[5].toLowerCase().startsWith("&h"))
          b = Number.parseInt(matches[5].substr(2), 16) / 0xff;
        else 
          b = Number.parseInt(matches[5]) / 0xff;
        
        color = new vscode.Color(r, g, b, 1);
      }

      const co = new vscode.ColorInformation(range, color);
      array.push(co);
    }
    return array;
  }

  public provideColorPresentations(color: vscode.Color, context: { document: vscode.TextDocument, range: vscode.Range }): vscode.ColorPresentation[] {
    const array = new Array<vscode.ColorPresentation>();
    const cp = new vscode.ColorPresentation(`RGB(${color.red * 0xff}, ${color.green * 0xff}, ${color.blue * 0xff})`);
    cp.textEdit = new vscode.TextEdit(context.range, `RGB(${color.red * 0xff}, ${color.green * 0xff}, ${color.blue * 0xff})`);

    array.push(cp);

    return array;
  }
}

export default vscode.languages.registerColorProvider(
  { scheme: "file", language: "vbs" },
  new VBSColorProvider()
);
