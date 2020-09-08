import { languages, Location, TextDocument, Position } from 'vscode';
import { VBS_MODE } from './util';

const AutoItDefinitionProvider = {
  provideDefinition(document : TextDocument, position : Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    let defRegex = new RegExp(`(?:Function|Sub)\\s${lookup}\\(`, "i");
    let found = docText.match(defRegex);
    if (found) {
      return new Location(document.uri, document.positionAt(found.index));
    }

    defRegex = new RegExp(`(?:Dim|Const)\\s${lookup}\\s?=?`, 'i');
    found = docText.match(defRegex);
    if (found) {
      return new Location(document.uri, document.positionAt(found.index));
    }

    return null;
  },
};

const defProvider = languages.registerDefinitionProvider(VBS_MODE, AutoItDefinitionProvider);

module.exports = defProvider;
