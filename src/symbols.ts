import { languages, SymbolKind, TextLine, Location, DocumentSymbol } from 'vscode';
import UTILS from './util';
import PATTERNS from './patterns';

const isSkippableLine = (line: TextLine) => {
  const skipChars = ["'"];

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
    const result : DocumentSymbol[] = [];
    const found = [];
    let name: string;

    const lastIdent = new Array<string>();
    lastIdent.push('');
  
    const VAR = RegExp(PATTERNS.VAR.source, 'i');
    const FUNCTION = RegExp(PATTERNS.FUNCTION.source, 'i');
    const CLASS = RegExp(PATTERNS.CLASS.source, 'i');
    const PROP = RegExp(PATTERNS.PROP.source, 'i');
    
    // Get the number of lines in the document to loop through
    const lineCount = Math.min(doc.lineCount, 10000);
    for (let lineNum = 0; lineNum < lineCount; lineNum++) {
      const line = doc.lineAt(lineNum);
      const lineText = line.text;
      const pos = new Location(doc.uri, line.range);

      if (isSkippableLine(line)) {
        // eslint-disable-next-line no-continue
        continue;
      }
      
      let matches = FUNCTION.exec(lineText);
      if (matches) {
        name = matches[2];
        let symKind = SymbolKind.Function;
        if (matches[1].toLowerCase() === "sub")
          if (name.toLowerCase() == "class_initialize" || name.toLowerCase() == "class_terminate")
            symKind = SymbolKind.Constructor;
          else
            symKind = SymbolKind.Method;

        const functionSymbol  = new DocumentSymbol(name, '', symKind, line.range, line.range);
        
        result.push(functionSymbol);
        found.push(name);
        lastIdent.push(name);
      }

      matches = CLASS.exec(lineText);
      if (matches) {
        name = matches[1];
        const classSymbol  = new DocumentSymbol(name, '', SymbolKind.Class, line.range, line.range);
        result.push(classSymbol);
        found.push(name);
        lastIdent.push(name);
      }

      matches = PROP.exec(lineText);
      if (matches) {
        name = matches[1];
        const classSymbol  = new DocumentSymbol(name, '', SymbolKind.Property, line.range, line.range);
        result.push(classSymbol);
        found.push(name);
        lastIdent.push(name);
      }
      
      matches = VAR.exec(lineText);
      if (matches) {
        let name = matches[2];
        let symKind = SymbolKind.Variable;
        if (matches[1].toLowerCase() === "const")
          symKind = SymbolKind.Constant;

        const variableSymbol  = new DocumentSymbol(name, '', symKind, line.range, line.range);
        result.push(variableSymbol);
        found.push(name);

      }
      // }
    }

    return result;
  },
});
