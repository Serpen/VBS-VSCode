import { languages, Hover, TextDocument, Position } from 'vscode';
import definitions from './definitions';
import PATTERNS from './patterns';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';

    if (word.trim() == "")
      return;
    
    let hover = definitions.find(key => key.label.toLowerCase() == word.toLowerCase())
    if (hover) {
      console.log(word + " | " + hover);
      return new Hover([hover.detail, hover.documentation]);
    }
    
    const text = document.getText();

    let matches = PATTERNS.DEF(word).exec(text);
    if (matches) {
      console.log(word + " | " + matches);
      return new Hover(matches);
    }
  },
});