import { languages, Location, TextDocument, Position, Uri } from 'vscode';
import { GlobalSourceImport, GlobalSourceImportFile, ObjectSourceImport, ObjectSourceImportFile, SourceImportFiles, SourceImports as SourceImports } from './extension';
import * as PATTERNS from './patterns';

export default languages.registerDefinitionProvider({ scheme: 'file', language: 'vbs' }, {
  provideDefinition(document: TextDocument, position: Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    const posLoc: Location[] = [];

    let match = PATTERNS.DEF(docText, lookup);
    if (match)
      posLoc.push(new Location(document.uri, document.positionAt(match.index)));

    match = PATTERNS.DEFVAR(docText, lookup);
    if (match)
      posLoc.push(new Location(document.uri, document.positionAt(match.index)));

    posLoc.push(...findExtDef(GlobalSourceImport, lookup, Uri.file(GlobalSourceImportFile)))
    posLoc.push(...findExtDef(ObjectSourceImport, lookup, Uri.file(ObjectSourceImportFile)))

    for (let index = 0; index < SourceImports.length; index++)
      posLoc.push(...findExtDef(SourceImports[index], lookup, Uri.file(SourceImportFiles[index])))

    if (posLoc.length > 0)
      return posLoc[0];
  }
},
);

function findExtDef(docText: string, lookup: string, docuri: Uri) : Location[] {
  const posloc : Location[] = [];
  let match = PATTERNS.DEF(docText, lookup);
  if (match) {
    const line = docText.slice(0, match.index).match(/\n/g).length;
    posloc.push(new Location(docuri, new Position(line, 0)));
  }

  match = PATTERNS.DEFVAR(docText, lookup);
  if (match) {
    const line = docText.slice(0, match.index).match(/\n/g).length;
    posloc.push(new Location(docuri, new Position(line, 0)));
  }
  return posloc;
}
