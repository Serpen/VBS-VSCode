import { languages, SymbolKind, TextLine, DocumentSymbol, Range, workspace } from 'vscode';
import PATTERNS from './patterns';

export default languages.registerDocumentSymbolProvider({ scheme: 'file', language: 'vbs' }, {
  provideDocumentSymbols(doc) {
    const result: DocumentSymbol[] = [];

    const FUNCTION = RegExp(PATTERNS.FUNCTION.source, 'i');
    const CLASS = RegExp(PATTERNS.CLASS.source, 'i');
    const PROP = RegExp(PATTERNS.PROP.source, 'i');
    const endLine = (/(?:^|:)[\t ]*End\s+(Sub|Class|Function|Property)/i);

    const showVariableSymbols: boolean = workspace.getConfiguration("vbs").get("showVariableSymbols");

    let currentBlock: DocumentSymbol[] = [];
    let waitCurrentBlockEnd: String[] = [];

    // Get the number of lines in the document to loop through
    const lineCount = Math.min(doc.lineCount, 10000);
    for (let lineNum = 0; lineNum < lineCount; lineNum++) {
      const line = doc.lineAt(lineNum);

      if (line.isEmptyOrWhitespace || line.text.charAt(line.firstNonWhitespaceCharacterIndex) == "'")
        continue;

      let name: string;
      let symbol: DocumentSymbol;

      let matches: RegExpMatchArray = [];

      if ((matches = CLASS.exec(line.text)) !== null) {
        name = matches[1];
        symbol = new DocumentSymbol(name, '', SymbolKind.Class, line.range, line.range);
        waitCurrentBlockEnd.push("class")

      } else if ((matches = FUNCTION.exec(line.text)) !== null) {
        name = matches[2];
        let detail: string = "";
        let symKind = SymbolKind.Function;
        if (matches[1].toLowerCase() === "sub")
          if (name.toLowerCase() == "class_initialize()" || name.toLowerCase() == "class_terminate()") {
            symKind = SymbolKind.Constructor;
            waitCurrentBlockEnd.push("sub");
          } else {
            detail = "Sub";
            waitCurrentBlockEnd.push(detail.toLowerCase());
          }
        else {
          detail = "Function";
          waitCurrentBlockEnd.push(detail.toLowerCase());
        }

        symbol = new DocumentSymbol(name, detail, symKind, line.range, line.range);
      } else if ((matches = PROP.exec(line.text)) !== null) {
        name = matches[2];
        symbol = new DocumentSymbol(name, matches[1], SymbolKind.Property, line.range, line.range);
        waitCurrentBlockEnd.push("property");
      } else if (showVariableSymbols) {
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

      if ((matches = endLine.exec(line.text)) !== null)
        if (waitCurrentBlockEnd[waitCurrentBlockEnd.length - 1] == matches[1].toLowerCase()) {
          currentBlock.pop();
          waitCurrentBlockEnd.pop();
        } else {
          currentBlock.pop();
          console.log("symbol wrong ending (awaiting closing for " + waitCurrentBlockEnd.pop().toString() + " got " + matches[1].toLowerCase() + ") in " + doc.uri + " " + line.lineNumber)
        }


    }

    if (waitCurrentBlockEnd.length > 0)
      console.log(waitCurrentBlockEnd);
    if (currentBlock.length > 0)
      console.log(currentBlock);

    return result;
  },
});

