import fs from 'fs';
import { languages, Location, TextDocument, Position, workspace, Uri } from 'vscode';
import PATTERNS from './patterns';

export default languages.registerDefinitionProvider({ scheme: 'file', language: 'vbs' }, {
  provideDefinition(document: TextDocument, position: Position) {
    const lookupRange = document.getWordRangeAtPosition(position);
    const lookup = document.getText(lookupRange);
    const docText = document.getText();

    let found = docText.match(PATTERNS.DEF(lookup));
    if (found)
      return new Location(document.uri, document.positionAt(found.index!));

    const ExtraDocument: string = workspace.getConfiguration("vbs").get("includes");
    if (ExtraDocument != '' && fs.statSync(ExtraDocument)) {
      const ExtraDocumentText = fs.readFileSync(ExtraDocument).toString();
      found = ExtraDocumentText.match(PATTERNS.DEF(lookup));
      if (found)
        return new Location(Uri.file(ExtraDocument), document.positionAt(found.index!));
    }

  },

});
