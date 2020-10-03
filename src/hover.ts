import { languages, Hover, TextDocument, Position, workspace } from 'vscode';
import PATTERNS from './patterns';
import { GlobalSourceImport, SourceImports } from './extension';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';

    if (word.trim() == "")
      return;

    const text = document.getText();

    let matches = PATTERNS.DEF(word).exec(text);
    if (matches)
      return new Hover(matches);
    
    const incs = [GlobalSourceImport, ...SourceImports];
    for (const ExtraDocText of incs) {
      matches = PATTERNS.DEF(word).exec(ExtraDocText);

      if (matches)
        return new Hover(matches[1])
    }
  },
});