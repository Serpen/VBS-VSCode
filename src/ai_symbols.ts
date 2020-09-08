import { languages, SymbolInformation, SymbolKind, workspace, TextLine } from 'vscode';
import UTILS from './util';
import PATTERNS from './patterns';

const variablePattern = /(?:Dim|Const)\s(\w+)/g;
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
      const { text } = line;

      if (isSkippableLine(line)) {
        // eslint-disable-next-line no-continue
        continue;
      }

      funcName = PATTERNS.FUNCTION.exec(text);
      if (funcName && !found.includes(funcName[1])) {
        const functionSymbol = new SymbolInformation(funcName[1], SymbolKind.Function, line.range);
        result.push(functionSymbol);
        found.push(funcName[1]);
      }

      if (config.showVariablesInGoToSymbol) {
        const variables = text.match(variablePattern);
        console.log("1");
        if (!variables) {
          // eslint-disable-next-line no-continue
          continue;
        }
        console.log("2");
        
        variables.forEach(variable => {
          if (found.includes(variable)) {
            return;
          }
          console.log("3");

          const variableSymbol = new SymbolInformation(variable, SymbolKind.Variable, line.range);
          result.push(variableSymbol);
          found.push(variable);
        });
      }
    }

    return result;
  },
});
