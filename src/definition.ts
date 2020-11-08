import { languages, Location, TextDocument, Position, Uri, Range } from "vscode";
import { getImportsWithLocal } from "./Includes";
import * as PATTERNS from "./patterns";

function findExtDef(docText: string, lookup: string, docuri: Uri): Location[] {
  const posloc: Location[] = [];
  let match = PATTERNS.DEF(docText, lookup);
  if (match) {
    const pos = match.index + match[1].length;
    const line = docText.slice(0, pos).match(/\n/g).length;
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
    matches[6]?.split(",").filter(p => p.trim() === lookup).forEach(() => {
      const line = docText.slice(0, matches.index).match(/\n/g).length;
      locs.push(new Location(thisUri, new Position(line, 0)));
    });

  // last result should be nearest hit
  if (locs.length > 0)
    return [locs[locs.length - 1]];
  else
    return [];
}

function provideDefinition(doc: TextDocument, position: Position): Location[] {
  const lookupRange = doc.getWordRangeAtPosition(position);
  const lookup = doc.getText(lookupRange);
  const docText = doc.getText();

  const posLoc: Location[] = [];

  let match = PATTERNS.DEF(docText, lookup);
  if (match)
    posLoc.push(new Location(doc.uri, doc.positionAt(match.index)));

  match = PATTERNS.DEFVAR(docText, lookup);
  if (match)
    posLoc.push(new Location(doc.uri, doc.positionAt(match.index)));

  for (const item of getImportsWithLocal(doc))
    posLoc.push(...findExtDef(item[1].Content, lookup, item[1].Uri));

  // def for param must be above
  posLoc.push(...GetParamDef(doc.getText(new Range(new Position(0, 0), new Position(position.line + 1, 0))), lookup, doc.uri));

  return posLoc;
}

export default languages.registerDefinitionProvider({ scheme: "file", language: "vbs" }, { provideDefinition });
