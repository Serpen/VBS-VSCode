import { languages, CompletionItem, CompletionItemKind, Range, TextDocument, Position } from 'vscode';
import definitions from './definitions';
import { GlobalSourceImport, SourceImports } from './extension';
import PATTERNS from './patterns';

function getVariableCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = {};

  let matches : RegExpMatchArray;
  while ((matches = PATTERNS.VAR_COMMENT.exec(text)) !== null) {
    const name = matches[3];

    if (!(name in foundVals)) {
      let itmKind = CompletionItemKind.Variable;
      if (matches[1].toLowerCase().indexOf("const") > 0)
        itmKind = CompletionItemKind.Constant;

      const ci = new CompletionItem(name, itmKind);
      ci.documentation = matches[4];

      ci.detail = `[${scope}] ` + matches[1];

      foundVals[name] = true;
      CIs.push(ci);
    }
  }

  return CIs;
}

function getFunctionCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundFunctions = {};

  let matches : RegExpExecArray;
  while ((matches = PATTERNS.FUNCTION_COMMENT.exec(text)) !== null) {
    const functionName = matches[5];

    if (!(functionName in foundFunctions)) {
      let itmKind = CompletionItemKind.Function;
      if (matches[3].toLowerCase() == "sub")
        itmKind = CompletionItemKind.Method;
      const ci = new CompletionItem(functionName, itmKind);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = `[${scope}] ` + matches[2];

      foundFunctions[functionName] = true;
      CIs.push(ci);
    }
  }

  return CIs;
}

function getPropertyCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = {};

  let matches : RegExpMatchArray;
  while ((matches = PATTERNS.PROP_COMMENT.exec(text)) !== null) {
    const name = matches[3];

    if (!(name in foundVals)) {
      const ci = new CompletionItem(name, CompletionItemKind.Property);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = `[${scope}] ` + matches[2];

      foundVals[name] = true;
      CIs.push(ci);
    }
  }

  return CIs;
}

function getClassCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = {};

  let matches = PATTERNS.CLASS.exec(text);
  while (matches) {
    const name = matches[1];
    if (!(name in foundVals)) {
      foundVals[name] = true;
      let ci = new CompletionItem(name, CompletionItemKind.Class);
      ci.detail = `[${scope}] ` + name
      CIs.push(ci);
    }
    matches = PATTERNS.CLASS.exec(text);
  }

  return CIs;
}

function getCompletions(text: string, scope: string) {
  return [...getVariableCompletions(text, scope), ...getFunctionCompletions(text, scope), ...getPropertyCompletions(text, scope), ...getClassCompletions(text, scope)];

}

function provideCompletionItems(document: TextDocument, position: Position): CompletionItem[] {
  // Gather the functions created by the user
  const text = document.getText();
  let range = document.getWordRangeAtPosition(position);

  if (!range)
    range = new Range(position, position);

  // Remove completion offerings from commented lines
  const line = document.lineAt(position);
  if (line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
    return [];

  const VAR = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class|Property [GLT]et))[\t ]+/i; //fix: should again after var name
  if (VAR.test(line.text))
    return [];

  const retCI: CompletionItem[] = [];

  retCI.push(...getCompletions(text, "Local"));

  retCI.push(...getCompletions(GlobalSourceImport, "Global"));

  SourceImports.forEach(ImportDoc => {
    retCI.push(...getCompletions(ImportDoc, "Import"));
  });

  return [...definitions, ...retCI];
}

export default languages.registerCompletionItemProvider(
  { scheme: 'file', language: 'vbs' },
  { provideCompletionItems }
);
