import { languages, Location, TextDocument, Position, Uri } from 'vscode';
import { GlobalSourceImport, GlobalSourceImportFile, ObjectSourceImport, ObjectSourceImportFile, SourceImportFiles, SourceImports as SourceImports } from './extension';
import * as PATTERNS from './patterns';

export default languages.registerDefinitionProvider({ scheme: 'file', language: 'vbs' }, {
  provideDefinition(document: TextDocument, position: Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    let match = PATTERNS.DEF(docText, lookup);
    if (match)
      return new Location(document.uri, document.positionAt(match.index!));

    match = PATTERNS.DEF(GlobalSourceImport, lookup);
    if (match) {
      const line = GlobalSourceImport.slice(0, match.index).match(/\n/g)!.length;
      return new Location(Uri.file(GlobalSourceImportFile), new Position(line, 0));
    }

    match = PATTERNS.DEF(ObjectSourceImport, lookup);
    if (match) {
      const line = ObjectSourceImport.slice(0, match.index).match(/\n/g)!.length;
      return new Location(Uri.file(ObjectSourceImportFile), new Position(line, 0));
    }

    for (let index = 0; index < SourceImports.length; index++) {
      match = PATTERNS.DEF(SourceImports[index], lookup);

      if (match) {
        const line = SourceImports[index].slice(0, match.index).match(/\n/g)!.length;
        return new Location(Uri.file(SourceImportFiles[index]), new Position(line, 0));
      }
    }
  }

},

);
