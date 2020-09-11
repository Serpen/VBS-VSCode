import { languages, Location, TextDocument, Position } from 'vscode';
import UTILS from './util';
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

    return null;
  },
};

const defProvider = languages.registerDefinitionProvider(UTILS.VBS_MODE, DefinitionProvider);

module.exports = defProvider;
