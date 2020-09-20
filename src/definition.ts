import { languages, Location, TextDocument, Position } from 'vscode';
import PATTERNS from './patterns';

const DefinitionProvider = {
  provideDefinition(document: TextDocument, position: Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    let found = docText.match(PATTERNS.DEF(lookup));
    if (found)
      return new Location(document.uri, document.positionAt(found.index!));
    else
      return null;
  },
};

export default languages.registerDefinitionProvider({ scheme: 'file', language: 'vbs' }, DefinitionProvider);
