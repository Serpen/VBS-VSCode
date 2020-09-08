import { languages, SymbolInformation, SymbolKind, Location, Position, workspace, Uri } from 'vscode';
import fs from 'fs';

import PATTERNS from './patterns';

const variablePattern = /(\$\w+)/g;
const config = workspace.getConfiguration('vbs');

const makeSymbol = (name : string, type : SymbolKind, filePath : Uri, docLine : number) => {
  return new SymbolInformation(name, type, '', new Location(filePath, new Position(docLine, 0)));
};

async function provideWorkspaceSymbols(search : string) {
  const symbols : SymbolInformation[] = [];

  // Don't start searching when it's empty
  if (!search) {
    return [];
  }

  // Enable searching for leading $
  const searchFilter = new RegExp(search.replace('$', '\\$'), 'i');

  // Get list of AutoIt files in workspace
  await workspace.findFiles('**/*.vbs').then(data => {
    data.forEach(file => {
      const foundVars = [];

      fs.readFileSync(file.fsPath)
        .toString()
        .split('\n')
        .forEach((line, index) => {
          let symbolKind : SymbolKind;
          const variableFound = variablePattern.exec(line);
          const functionFound = PATTERNS.FUNCTION.exec(line);

          if (variableFound && config.showVariablesInGoToSymbol) {
            const { 1: newName } = variableFound;

            // Filter based on search (if it's not empty)
            if (!searchFilter.exec(newName)) {
              return false;
            }
            symbolKind = SymbolKind.Variable;

            if (foundVars.indexOf(newName) === -1) {
              foundVars.push(newName);
              return symbols.push(makeSymbol(newName, symbolKind, file, index));
            }
            return false;
          }

          if (functionFound) {
            const { 1: newName } = functionFound;
            if (!searchFilter.exec(newName)) {
              return false;
            }
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
