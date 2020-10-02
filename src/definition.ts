import fs from 'fs';
import { languages, Location, TextDocument, Position, workspace, Uri } from 'vscode';
import PATTERNS from './patterns';

export default languages.registerDefinitionProvider({ scheme: 'file', language: 'vbs' }, {
  provideDefinition(document: TextDocument, position: Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    let match = docText.match(PATTERNS.DEF(lookup));
    if (match)
      return new Location(document.uri, document.positionAt(match.index!));

    const ExtraDocument: string = workspace.getConfiguration("vbs").get("includes");
    if (ExtraDocument != '' && fs.statSync(ExtraDocument)) {
      const ExtraDocumentText = fs.readFileSync(ExtraDocument).toString();
      match = ExtraDocumentText.match(PATTERNS.DEF(lookup));

      const line = ExtraDocumentText.slice(0, match.index).match(/\n/g).length;

      if (match)
        return new Location(Uri.file(ExtraDocument), new Position(line, 0));
    }

  },

});
