import { languages, Hover, TextDocument, Position, Range } from "vscode";
import * as PATTERNS from "./patterns";
import { getImportsWithLocal } from "./Includes";

function GetHover(docText: string, lookup: string, scope: string): Hover[] {
  const results: Hover[] = [];
  let matches = PATTERNS.DEF(docText, lookup);
  if (matches) {
    if (matches[1]) {
      const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
      if (summary[1])
        results.push(new Hover({ language: "vbs", value: `${matches[2]} ' [${scope}]\n' ${summary[1]}` }));
      else
        results.push(new Hover({ language: "vbs", value: `${matches[2]} ' [${scope}]` }));
    } else
      results.push(new Hover({ language: "vbs", value: `${matches[2]} ' [${scope}]` }));
  }

  matches = PATTERNS.DEFVAR(docText, lookup);
  if (matches) {
    if (matches[1]) {
      const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
      if (summary[1])
        results.push(new Hover({ language: "vbs", value: `${matches[2]} ' [${scope}]\n' ${summary[1]}` }));
      else
        results.push(new Hover({ language: "vbs", value: `${matches[2]} ' [${scope}]` }));
    } else
      results.push(new Hover({ language: "vbs", value: `${matches[2]} ' [${scope}]` }));
  }

  return results;
}

function GetParamHover(text: string, lookup: string): Hover[] {
  const hovers: Hover[] = [];

  let matches: RegExpExecArray;
  while (matches = PATTERNS.FUNCTION.exec(text))
    matches[6]?.split(",").filter(p => p.trim() === lookup).forEach(() => {
      hovers.push(new Hover({ language: "vbs", value: `${lookup} ' [Parameter]` }));
    });

  // last result should be nearest hit
  if (hovers.length > 0)
    return [hovers[hovers.length - 1]];
  else
    return [];
}

function provideHover(doc: TextDocument, position: Position): Hover {
  const wordRange = doc.getWordRangeAtPosition(position);
  const word: string = wordRange ? doc.getText(wordRange) : "";
  const line = doc.lineAt(position).text;

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
  hoverresults.push(...GetHover(doc.getText(), word, "Local"));

  for (const ExtraDocText of getImportsWithLocal(doc))
    hoverresults.push(...GetHover(ExtraDocText[1].Content, word, ExtraDocText[0]));

  // hoverresult for param must be above
  hoverresults.push(...GetParamHover(doc.getText(new Range(new Position(0, 0), new Position(position.line + 1, 0))), word));

  if (hoverresults.length > 0)
    return hoverresults[0];
  else
    return null;
}

export default languages.registerHoverProvider(
  { scheme: "file", language: "vbs" },
  { provideHover }
);
