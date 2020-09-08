import { languages, SymbolInformation, SymbolKind, workspace, TextLine } from 'vscode';
import UTILS from './util';
import PATTERNS from './patterns';

const config = workspace.getConfiguration('vbs');

const isSkippableLine = (line: TextLine) => {
  const skipChars = [';', '#'];

  if (line.isEmptyOrWhitespace) {
    return true;
  }

  const firstChar = line.text.charAt(line.firstNonWhitespaceCharacterIndex);
  if (skipChars.includes(firstChar)) {
    return true;
  }

  return false;
};


module.exports = languages.registerDocumentSymbolProvider(UTILS.VBS_MODE, {
  provideDocumentSymbols(doc) {
    const result = [];
    const found = [];
    let matches: RegExpExecArray;
    let name: string;

    // Get the number of lines in the document to loop through
    const lineCount = Math.min(doc.lineCount, 10000);
    for (let lineNum = 0; lineNum < lineCount; lineNum += 1) {
      const line = doc.lineAt(lineNum);

      if (isSkippableLine(line)) {
        // eslint-disable-next-line no-continue
        continue;
      }

      let matches = PATTERNS.FUNCTION.exec(line.text);
      if (matches) {
        name = matches[2];
        let symKind = SymbolKind.Function;
        if (matches[1].toLowerCase() === "sub")
          if (name.toLowerCase() == "class_initialize" || name.toLowerCase() == "class_terminate")
            symKind = SymbolKind.Constructor;
          else
            symKind = SymbolKind.Method;

        const functionSymbol = new SymbolInformation(name, symKind, line.range);
        result.push(functionSymbol);
        found.push(name);
      }

      matches = PATTERNS.CLASS.exec(line.text);
      if (matches) {
        name = matches[1];
        const classSymbol = new SymbolInformation(name, SymbolKind.Class, line.range);
        result.push(classSymbol);
        found.push(name);
      }

      matches = PATTERNS.PROP.exec(line.text);
      if (matches) {
        name = matches[1];
        const classSymbol = new SymbolInformation(name, SymbolKind.Property, line.range);
        result.push(classSymbol);
        found.push(name);
      }

      // if (!config.showVariablesInGoToSymbol) {
      matches = PATTERNS.VAR.exec(line.text);
      if (matches) {
        let name = matches[2];
        let symKind = SymbolKind.Variable;
        if (matches[1].toLowerCase() === "const")
          symKind = SymbolKind.Constant;

        const variableSymbol = new SymbolInformation(name, symKind, line.range);
        result.push(variableSymbol);
        found.push(name);
      }
      // }
    }

    return result;
  },
});
