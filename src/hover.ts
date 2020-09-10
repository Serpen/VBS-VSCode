"use strict";
import { languages, Hover, CompletionItem, TextDocument, Position } from 'vscode';
import vbs from './definitions/index';
import UTILS from './util';
import PATTERNS from './patterns';

module.exports = languages.registerHoverProvider(UTILS.VBS_MODE, {
  provideHover(document : TextDocument, position : Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word : string = wordRange ? document.getText(wordRange) : '';

    let hover = vbs.find(key => key.label.toLowerCase() == word.toLowerCase())
    if (hover)
      return new Hover([hover.detail, hover.documentation]);

    const text = document.getText();

    let matches = PATTERNS.Function(word).exec(text);
    if (matches) {
      return new Hover(matches[1] + " " + word);
    }
  },
});