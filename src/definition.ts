import { languages, Location, TextDocument, Position, Uri, Range } from 'vscode';
import { Includes } from './extension';
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

    for (const item of Includes)
      posLoc.push(...findExtDef(item[1].Content, lookup, item[1].Uri))

    // def for param must be above
    posLoc.push(...GetParamDef(document.getText(new Range(new Position(0, 0), new Position(position.line + 1, 0))), lookup, document.uri));

    if (posLoc.length > 0)
      return posLoc[0];
  }
},
);

function findExtDef(docText: string, lookup: string, docuri: Uri): Location[] {
  const posloc: Location[] = [];
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

function GetParamDef(docText: string, lookup: string, thisUri: Uri): Location[] {
  const locs: Location[] = [];

  let matches: RegExpExecArray;
  while (matches = PATTERNS.FUNCTION.exec(docText))
    matches[6]?.split(",").filter(p => p.trim() === lookup).forEach(param => {
      const line = docText.slice(0, matches.index).match(/\n/g).length;
      locs.push(new Location(thisUri, new Position(line, 0)));
    });

  // last result should be nearest hit
  if (locs.length > 0)
    return [locs[locs.length - 1]];
  else
    return [];
}