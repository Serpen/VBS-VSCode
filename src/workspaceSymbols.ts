import { languages, SymbolInformation, SymbolKind, Location, Position, workspace, Uri } from 'vscode';
import fs from 'fs';

import PATTERNS from './patterns';

const config = workspace.getConfiguration('vbs');

function makeSymbol(name: string, type: SymbolKind, filePath: Uri, docLine: number) {
  return new SymbolInformation(name, type, '', new Location(filePath, new Position(docLine, 0)));
}

async function provideWorkspaceSymbols(search: string) {
  const symbols: SymbolInformation[] = [];

  // Don't start searching when it's empty
  if (!search) {
    return [];
  }

  // Get list of vbs files in workspace
  await workspace.findFiles('**/*.vbs').then(data => {
    data.forEach(file => {
      const foundVars = Array<string>();

      const VAR = RegExp(PATTERNS.VAR.source, 'i');
      const FUNCTION = RegExp(PATTERNS.FUNCTION.source, 'i');
      const CLASS = RegExp(PATTERNS.CLASS.source, 'i');
      const PROP = RegExp(PATTERNS.PROP.source, 'i');

      fs.readFileSync(file.fsPath)
        .toString()
        .split('\n')
        .forEach((line, index) => {
          let symbolKind: SymbolKind;
          const variableFound = VAR.exec(line);
          const functionFound = FUNCTION.exec(line);

          if (variableFound && config.showVariablesInGoToSymbol) {
            const { 1: newName } = variableFound;

            symbolKind = SymbolKind.Variable;

            if (foundVars.indexOf(newName) === -1) {
              foundVars.push(newName);
              return symbols.push(makeSymbol(newName, symbolKind, file, index));
            }
            return false;
          }

          if (functionFound) {
            const { 3: newName } = functionFound;
            symbolKind = SymbolKind.Function;
            return symbols.push(makeSymbol(newName, symbolKind, file, index));
          }
          return false;
        });
    });
  });
  return symbols;
}

const workspaceSymbolProvider = languages.registerWorkspaceSymbolProvider({
  provideWorkspaceSymbols,
});

export default workspaceSymbolProvider;
