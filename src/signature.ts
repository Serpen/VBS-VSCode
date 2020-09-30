import fs from 'fs';
import { languages, SignatureHelp, SignatureInformation, ParameterInformation, MarkdownString, TextDocument, Position, workspace, CompletionItem } from 'vscode';
import defaultSigs from './definitions/functions.json';
import PATTERNS from './patterns';

/**
 * Reduces a partial line of code to the current Function for parsing
 * @param {string} code The line of code
 */
function getParsableCode(code: string): string {
  const reducedCode = code
    .replace(/\w+\([^()]*\)/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(/"[^"]*(?=$)/g, '') // Remove double quote and text at end of line
    .replace(/\([^()]*\)/g, '') // Remove paren sets
    .replace(/\({2,}/g, '('); // Reduce multiple open parens

  return reducedCode;
}

function getCurrentFunction(code: string) {
  const parenSplit = code.split('(');
  let index: number;
  if (parenSplit.length == 1)
    index = 0;
  else
    index = parenSplit.length - 2;

  // Get the 2nd to last item (right in front of last open paren)
  // and clean up the results
  return parenSplit[index].match(/(.*)\b(\w+)/)[2];
}

function countCommas(code: string) {
  // Find the position of the closest/last open paren
  const openParen = code.lastIndexOf('(');
  // Count non-string commas in text following open paren
  let commas = code.slice(openParen).match(/(?!\B["'][^"']*),(?![^"']*['"]\B)/g);
  if (commas === null) {
    return 0;
  } else {
    return commas.length;
  }
}

function getCallInfo(doc: TextDocument, pos: Position) {
  // Acquire the text up the point where the current cursor/paren/comma is at
  const codeAtPosition = doc.lineAt(pos).text.substring(0, pos.character);
  const cleanCode = getParsableCode(codeAtPosition);

  return {
    func: getCurrentFunction(cleanCode),
    commas: countCommas(cleanCode),
  };
}

function getSignatures(text: string, docComment: string): Map<string, SignatureInformation[]> {
  let map = new Map<string, SignatureInformation[]>();

  let matches: RegExpExecArray;
  while ((matches = PATTERNS.FUNCTION.exec(text)) !== null) {
    const si = new SignatureInformation(matches[2], docComment);
    matches[4].split(",").forEach(element => {
      si.parameters.push(new ParameterInformation(element.trim()));
    });

    let prevMatches;
    if ((prevMatches = map.get(matches[3])) !== undefined)
      map.set(matches[3], [si, ...prevMatches]);
    else
      map.set(matches[3], [si]);
  }

  return map;
}

export default languages.registerSignatureHelpProvider({ scheme: 'file', language: 'vbs' },
  {
    provideSignatureHelp(document, position) {
      const caller = getCallInfo(document, position);
      if (caller == null)
        return null;

      const ExtraDocument: string = workspace.getConfiguration("vbs").get("includes");
      let ExtraDocumentText: string = "";
      if (ExtraDocument != '' && fs.statSync(ExtraDocument))
        ExtraDocumentText = fs.readFileSync(ExtraDocument).toString();

      const sigs = new SignatureHelp();
      sigs.activeSignature = 0;
      sigs.activeParameter = caller.commas;

      let sig;
      if ((sig = getSignatures(document.getText(), "Local Function").get(caller.func)) !== undefined) {
        sigs.signatures.push(...sig);
      }
      if ((sig = getSignatures(ExtraDocumentText, "Included Function").get(caller.func)) !== undefined) {
        sigs.signatures.push(...sig);
      }

      return sigs;
    },
  },
  '(', ',', ' '
);
