import { languages, SignatureHelp, SignatureInformation, ParameterInformation,
  TextDocument, Position, SignatureHelpContext, CancellationToken } from "vscode";
import { getImportsWithLocal } from "./Includes";
import * as PATTERNS from "./patterns";

/**
 * Reduces a partial line of code to the current Function for parsing
 * @param {string} code The line of code
 */
function getParsableCode(code: string): string {
  const reducedCode = code
    .replace(/\w+\([^()]*\)/g, "")
    .replace(/"[^"]*"/g, "")
    .replace(/"[^"]*(?=$)/g, "") // Remove double quote and text at end of line
    .replace(/\([^()]*\)/g, "") // Remove paren sets
    .replace(/\({2,}/g, "("); // Reduce multiple open parens

  return reducedCode;
}

function getCurrentFunction(code: string) {
  const parenSplit = code.split("(");
  let index: number;
  if (parenSplit.length === 1)
    index = 0;
  else
    index = parenSplit.length - 2;

  // Get the 2nd to last item (right in front of last open paren)
  // and clean up the results
  return parenSplit[index].match(/(?:.*)\b(\w+)/)[1].toLowerCase();
}

function countCommas(code: string) {
  // Find the position of the closest/last open paren
  const openParen = code.lastIndexOf("(");
  // Count non-string commas in text following open paren
  const commas = code.slice(openParen).match(/(?!\B["'][^"']*),(?![^"']*['"]\B)/g);
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
    commas: countCommas(cleanCode)
  };
}

function getSignatures(text: string, docComment: string): Map<string, SignatureInformation[]> {
  const map = new Map<string, SignatureInformation[]>();
  let matches: RegExpExecArray | null;
  while ((matches = PATTERNS.FUNCTION.exec(text)) !== null) {
    const name = matches[5].toLowerCase();
    let documentation = "";

    if (matches[1]) {
      const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
      if (summary)
        documentation = summary[1];
      else
        documentation = docComment;
    }
    const si = new SignatureInformation(matches[4], documentation);
    matches[6]?.split(",").forEach(param => {
      const pi = new ParameterInformation(param.trim());
      if (matches[1]) {
        const paramComment = PATTERNS.PARAM_SUMMARY(matches[1], param.trim());
        if (paramComment)
          pi.documentation = paramComment[1];
      }
      si.parameters.push(pi);
    });


    let prevMatches: SignatureInformation[];
    if ((prevMatches = map.get(name)) !== undefined)
      map.set(name, [
        ...prevMatches,
        si
      ]);
    else
      map.set(name, [si]);
  }

  return map;
}

function provideSignatureHelp(doc: TextDocument, position: Position, _token: CancellationToken, context: SignatureHelpContext): SignatureHelp {
  const caller = getCallInfo(doc, position);
  if (caller === null)
    return null;

  const sighelp = new SignatureHelp();
  if (context.activeSignatureHelp)
    sighelp.activeSignature = context.activeSignatureHelp.activeSignature;
  else
    sighelp.activeSignature = 0;
  sighelp.activeParameter = caller.commas;

  let sig: SignatureInformation[] | undefined;
  if ((sig = getSignatures(doc.getText(), "Local").get(caller.func)) !== undefined) {
    sighelp.signatures.push(...sig.filter((sig2: SignatureInformation) => sig2.parameters.length >= caller.commas));
  }

  for (const item of getImportsWithLocal(doc)) {
    if ((sig = getSignatures(item[1].Content, item[0]).get(caller.func)) !== undefined) {
      sighelp.signatures.push(...sig.filter((sig2: SignatureInformation) => sig2.parameters.length >= caller.commas));
    }
  }

  return sighelp;
}

export default languages.registerSignatureHelpProvider(
  { language: "vbs" },
  { provideSignatureHelp }, "(", ",", " "
);
