import { languages, SymbolKind, TextLine, DocumentSymbol, Range, workspace } from 'vscode';
import PATTERNS from './patterns';

function isSkippableLine(line: TextLine) {
  const skipChars = ["'"];

  if (line.isEmptyOrWhitespace) {
    return true;
  }

  const firstChar = line.text.charAt(line.firstNonWhitespaceCharacterIndex);
  if (skipChars.includes(firstChar)) {
    return true;
  }

  return false;
}

export default languages.registerDocumentSymbolProvider({ scheme: 'file', language: 'vbs' }, {
  provideDocumentSymbols(doc) {
    const result: DocumentSymbol[] = [];

    const FUNCTION = RegExp(PATTERNS.FUNCTION.source, 'i');
    const CLASS = RegExp(PATTERNS.CLASS.source, 'i');
    const PROP = RegExp(PATTERNS.PROP.source, 'i');
    const endLine = (/(?:^|:)[\t ]*End\s+(?:Sub|Class|Function|Property)/i);

    const showVariableSymbols: boolean = workspace.getConfiguration("vbs").get("showVariableSymbols");

    let currentBlock: DocumentSymbol[] = [];
    let waitCurrentBlockEnd: String[] = [];

    // Get the number of lines in the document to loop through
    const lineCount = Math.min(doc.lineCount, 10000);
    for (let lineNum = 0; lineNum < lineCount; lineNum++) {
      const line = doc.lineAt(lineNum);

      if (isSkippableLine(line))
        // eslint-disable-next-line no-continue
        continue;

      let name: string;
      let symbol: DocumentSymbol;

      let matches: RegExpMatchArray = [];

      if ((matches = CLASS.exec(line.text)) !== null) {
        name = matches[1];
        symbol = new DocumentSymbol(name, '', SymbolKind.Class, line.range, line.range);
        waitCurrentBlockEnd.push("End Class")

      } else if ((matches = FUNCTION.exec(line.text)) !== null) {
        name = matches[2];
        let detail: string;
        let symKind = SymbolKind.Function;
        if (matches[1].toLowerCase() === "sub")
          if (name.toLowerCase() == "class_initialize()" || name.toLowerCase() == "class_terminate()")
            symKind = SymbolKind.Constructor;
          else
            detail = "Sub";
        else
          detail = "Function";

        symbol = new DocumentSymbol(name, detail, symKind, line.range, line.range);
        waitCurrentBlockEnd.push("End " + detail);

      } else if ((matches = PROP.exec(line.text)) !== null) {
        name = matches[2];
        symbol = new DocumentSymbol(name, matches[1], SymbolKind.Property, line.range, line.range);
      }
      
      if (showVariableSymbols) {
        while ((matches = PATTERNS.VAR2.exec(line.text)) !== null) {
          const varNames = matches[1].split(',');
          for (let i = 0; i < varNames.length; i++) {

            let name = varNames[i].trim();
            let symKind = SymbolKind.Variable;
            if (matches[1].toLowerCase().indexOf("const") > 0)
              symKind = SymbolKind.Constant;
            let r = new Range(line.lineNumber, 0, line.lineNumber, PATTERNS.VAR2.lastIndex);
            const variableSymbol = new DocumentSymbol(name, '', symKind, r, r);
            if (currentBlock.length == 0)
              result.push(variableSymbol);
            else
              currentBlock[currentBlock.length - 1].children.push(variableSymbol);
          }

        }
      }

      if (symbol != null) {
        if (currentBlock.length == 0)
          result.push(symbol);
        else
          currentBlock[currentBlock.length - 1].children.push(symbol);
        currentBlock.push(symbol);
      }

      if (endLine.test(line.text))
        currentBlock.pop();
    }
    return result;
  },
});

