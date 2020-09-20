import { languages, Location, TextDocument, Position } from 'vscode';
import PATTERNS from './patterns';

const DefinitionProvider = {
  provideDefinition(document: TextDocument, position: Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    let found = docText.match(PATTERNS.FUNC_DEF(lookup));
    if (found) {
      return new Location(document.uri, document.positionAt(found.index!));
    }

    found = docText.match(PATTERNS.VAR_DEF(lookup));
    if (found) {
      return new Location(document.uri, document.positionAt(found.index!));
    }

    found = docText.match(PATTERNS.PROP_DEF(lookup));
    if (found) {
      return new Location(document.uri, document.positionAt(found.index!));
    }

    found = docText.match(PATTERNS.CLASS_DEF(lookup));
    if (found) {
      return new Location(document.uri, document.positionAt(found.index!));
    }

    return null;
  },
};

export default languages.registerDefinitionProvider({ scheme: 'file', language: 'vbs' }, DefinitionProvider);
