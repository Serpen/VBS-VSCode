import { languages, Hover, TextDocument, Position } from 'vscode';
import * as PATTERNS from './patterns';
import { GlobalSourceImport, ObjectSourceImport, SourceImports } from './extension';

export default languages.registerHoverProvider({ scheme: 'file', language: 'vbs' }, {
  provideHover(document: TextDocument, position: Position) {
    const wordRange = document.getWordRangeAtPosition(position);
    const word: string = wordRange ? document.getText(wordRange) : '';
    const line = document.lineAt(position).text;

    const hoverresults: Hover[] = [];

    if (word.trim() === "")
      return null;

    if (!new RegExp(`^[^']*${word}`).test(line))
      return null;

    let count = 0;
    for (let i = 0; i < position.character; i++)
      if (line[i] === '"') count++;
    if (count % 2 === 1)
      return null;
    hoverresults.push(...GetHover(document.getText(), word, "[Local]"));
    hoverresults.push(...GetHover(GlobalSourceImport, word, "[Global]"));
    hoverresults.push(...GetHover(ObjectSourceImport, word, "[Global]"));

    for (const ExtraDocText of SourceImports)
      hoverresults.push(...GetHover(ExtraDocText, word, "[Import]"));

    if (hoverresults.length > 0)
      return hoverresults[0];
  },
});

function GetHover(docText: string, lookup: string, scope: string) : Hover[] {
  const results: Hover[] = [];
  let matches = PATTERNS.DEF(docText, lookup);
  if (matches) {
    if (matches[1]) {
      const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
      if (summary[1])
        results.push(new Hover({ language: "vbs", value: matches[2] + " ' " + scope + "\n' " + summary[1] }));
      else
        results.push(new Hover({ language: "vbs", value: matches[2] + " ' " + scope }));
    } else
      results.push(new Hover({ language: "vbs", value: matches[2] + " ' " + scope }));
  }

  matches = PATTERNS.DEFVAR(docText, lookup);
  if (matches) {
    if (matches[1]) {
      const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
      if (summary[1])
        results.push(new Hover({ language: "vbs", value: matches[2] + " ' " + scope + "\n' " + summary[1] }));
      else
        results.push(new Hover({ language: "vbs", value: matches[2] + " ' " + scope }));
    } else
      results.push(new Hover({ language: "vbs", value: matches[2] + " ' " + scope }));
  }
  return results;
}