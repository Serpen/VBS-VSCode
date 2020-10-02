import { languages, Hover, TextDocument, Position, workspace } from 'vscode';
import PATTERNS from './patterns';
import fs from 'fs';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';

    if (word.trim() == "")
      return;

    const text = document.getText();

    let matches = PATTERNS.DEF(word).exec(text);
    if (matches) {
      return new Hover(matches);
    }

    const ExtraDocument: string = workspace.getConfiguration("vbs").get("includes");
    if (ExtraDocument != '' && fs.statSync(ExtraDocument)) {
      const ExtraDocumentText = fs.readFileSync(ExtraDocument).toString();
      matches = PATTERNS.DEF(word).exec(ExtraDocumentText);

      if (matches)
        return new Hover(matches[1])
    }
  },
});