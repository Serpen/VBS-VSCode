import * as vscode from "vscode";
import * as PATTERNS from "./patterns";

class VBSColorProvider implements vscode.DocumentColorProvider {

  // eslint-disable-next-line no-magic-numbers
  private readonly MAX_COLOR = 0xff;

  public provideDocumentColors(doc: vscode.TextDocument): vscode.ColorInformation[] {
    const array = new Array<vscode.ColorInformation>();
    let matches: RegExpExecArray;
    while ((matches = PATTERNS.COLOR.exec(doc.getText())) !== null) {
      const pos = doc.positionAt(matches.index);
      const posEnd = doc.positionAt(matches.index + matches[0].length);
      const range = new vscode.Range(pos, posEnd);

      let color: vscode.Color;

      if (matches[1]) {
        switch (matches[1].toLowerCase()) {
        case "vbblack":
          color = new vscode.Color(0, 0, 0, 1);
          break;
        case "vbblue":
          color = new vscode.Color(0, 0, 1, 1);
          break;
        case "vbcyan":
          color = new vscode.Color(0, 1, 1, 1);
          break;
        case "vbgreen":
          color = new vscode.Color(0, 1, 0, 1);
          break;
        case "vbmagenta":
          color = new vscode.Color(1, 0, 1, 1);
          break;
        case "vbred":
          color = new vscode.Color(1, 0, 0, 1);
          break;
        case "vbwhite":
          color = new vscode.Color(1, 1, 1, 1);
          break;
        case "vbyellow":
          color = new vscode.Color(1, 1, 0, 1);
          break;
        }
      } else if (matches[2]) {
        let r: number, g: number, b: number;
        if (matches[3].toLowerCase().startsWith("&h"))
          r = Number.parseInt(matches[3].substr(2), 16) / this.MAX_COLOR;
        else
          r = Number.parseInt(matches[3]) / this.MAX_COLOR;

        if (matches[4].toLowerCase().startsWith("&h"))
          g = Number.parseInt(matches[4].substr(2), 16) / this.MAX_COLOR;
        else
          g = Number.parseInt(matches[4]) / this.MAX_COLOR;

        if (matches[5].toLowerCase().startsWith("&h"))
          b = Number.parseInt(matches[5].substr(2), 16) / this.MAX_COLOR;
        else
          b = Number.parseInt(matches[5]) / this.MAX_COLOR;

        color = new vscode.Color(r, g, b, 1);
      } else if (matches[6] && (/color/i).test(doc.lineAt(pos.line).text)) {
        const r = Number.parseInt(matches[6].substr(2, 2), 16) / this.MAX_COLOR;
        const b = Number.parseInt(matches[6].substr(4, 2), 16) / this.MAX_COLOR;
        const g = Number.parseInt(matches[6].substr(6, 2), 16) / this.MAX_COLOR;
        color = new vscode.Color(r, g, b, 1);
      }

      if (color)
        array.push(new vscode.ColorInformation(range, color));
    }

    return array;
  }

  public provideColorPresentations(color: vscode.Color): vscode.ColorPresentation[] {
    return [new vscode.ColorPresentation(`RGB(${color.red * this.MAX_COLOR}, ${color.green * this.MAX_COLOR}, ${color.blue * this.MAX_COLOR})`)];
  }
}

export default vscode.languages.registerColorProvider(
  { scheme: "file", language: "vbs" },
  new VBSColorProvider()
);
