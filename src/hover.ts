import { languages, Hover, TextDocument, Position } from 'vscode';
import definitions from './definitions';
import PATTERNS from './patterns';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';

    let hover = definitions.find(key => key.label.toLowerCase() == word.toLowerCase())
    if (hover)
      return new Hover([hover.detail, hover.documentation]);

    const text = document.getText();

    let matches = PATTERNS.Function(word).exec(text);
    if (matches) {
      return new Hover(matches[1] + " " + word);
    }
  },
});