import { languages, Hover, TextDocument, Position } from 'vscode';
import * as PATTERNS from './patterns';
import { GlobalSourceImport, ObjectSourceImport, SourceImports } from './extension';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';
    const line = document.lineAt(position).text;

    if (word.trim() === "")
      return null;

    if (!new RegExp(`^[^']*${word}`).test(line))
      return null;

    let count = 0;
    for (let i = 0; i < position.character; i++)
      if (line[i] === '"') count++;
    if (count % 2 === 1)
      return null;


    let matches = PATTERNS.DEF(document.getText(), word);
    if (matches)
      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        if (summary[1])
          return new Hover({ language: "vbs", value: matches[2] + " ' [Local]\n' " + summary[1] });
        else
          return new Hover({ language: "vbs", value: matches[2] + " ' [Local]" });
      } else
        return new Hover({ language: "vbs", value: matches[2] + " ' [Local]" });

    for (const ExtraDocText of [GlobalSourceImport, ObjectSourceImport, ...SourceImports]) {
      matches = PATTERNS.DEF(ExtraDocText, word);
      if (matches)
        if (matches[1]) {
          const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
          if (summary[1])
            return new Hover({ language: "vbs", value: matches[2] + " ' [Import/Global]\n' " + summary[1] });
          else
            return new Hover({ language: "vbs", value: matches[2] + " ' [Import/Global]" });
        } else
          return new Hover({ language: "vbs", value: matches[2] + " ' [Import/Global]" });
    }
  },
});