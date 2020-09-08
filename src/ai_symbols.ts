import { languages, SymbolInformation, SymbolKind, workspace, TextLine } from 'vscode';
import UTILS from './util';
import PATTERNS from './patterns';

const config = workspace.getConfiguration('vbs');

const isSkippableLine = (line : TextLine) => {
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
    let funcName : RegExpExecArray;

    // Get the number of lines in the document to loop through
    const lineCount = Math.min(doc.lineCount, 10000);
    for (let lineNum = 0; lineNum < lineCount; lineNum += 1) {
      const line = doc.lineAt(lineNum);

      if (isSkippableLine(line)) {
        // eslint-disable-next-line no-continue
        continue;
      }

      funcName = PATTERNS.FUNCTION.exec(line.text);
      if (funcName && !found.includes(funcName[2])) {
        let symKind = SymbolKind.Function;
        if (funcName[1].toLowerCase() === "sub")
          symKind = SymbolKind.Method;
          
        const functionSymbol = new SymbolInformation(funcName[2], symKind, line.range);
        result.push(functionSymbol);
        found.push(funcName[2]);
      }

      // if (!config.showVariablesInGoToSymbol) {
        let varName = PATTERNS.VAR.exec(line.text);
        if (varName && !found.includes(varName[2])) {
          let symKind = SymbolKind.Variable;
          if (varName[1].toLowerCase() === "const")
            symKind = SymbolKind.Constant;
          
          const variableSymbol = new SymbolInformation(varName[2], symKind, line.range);
          result.push(variableSymbol);
          found.push(varName[2]);

        }       
      // }
    }

    return result;
  },
});
