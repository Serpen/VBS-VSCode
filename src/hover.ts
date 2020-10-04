import { languages, Hover, TextDocument, Position, workspace } from 'vscode';
import PATTERNS from './patterns';
import { GlobalSourceImport, SourceImports } from './extension';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';
    const line = document.lineAt(position).text;

    if (word.trim() == "")
      return null;

    if (!new RegExp(`^[^']*${word}`).test(line))
      return null;


    let matches = PATTERNS.DEF(document.getText(), word);
    if (matches)
      if (matches[0].startsWith("\t"))
        return new Hover(matches[0] + "\n[Local]");
      else
        return new Hover("\t" + matches[0] + "\n[Local]"); // why??


    for (const ExtraDocText of [GlobalSourceImport, ...SourceImports]) {
      matches = PATTERNS.DEF(ExtraDocText, word);
      if (matches)
        if (matches[0].startsWith("\t"))
          return new Hover(matches[0] + "\n[Import/Global]");
        else
          return new Hover("\t" + matches[0] + "\n[Import/Global]"); // why??
    }
  },
});