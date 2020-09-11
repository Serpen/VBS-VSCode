import { languages, CompletionItem, CompletionItemKind, Range, TextDocument, Position } from 'vscode';
import definitions from './definitions';
import PATTERNS from './patterns';

function createNewCompletionItem(kind: CompletionItemKind, name: string, strDetail = '') {
  const compItem = new CompletionItem(name, kind);

  if (strDetail != '')
    compItem.detail = strDetail;
  else {
    switch (kind) {
      case CompletionItemKind.Constant:
        compItem.detail = "Document Constant";
        break;
      case CompletionItemKind.Function:
        compItem.detail = "Document Function";
        break;
      case CompletionItemKind.Method:
        compItem.detail = "Document Sub";
        break;
      case CompletionItemKind.Variable:
        compItem.detail = "Document Variable";
        break;
      case CompletionItemKind.Property:
        compItem.detail = "Document Property";
        break;
    }
  }
  return compItem;
}

function getVariableCompletions(text: string): CompletionItem[] {
  const variables: CompletionItem[] = [];
  const foundVariables = {};
  let variableName: string;

  let matches = PATTERNS.VAR_COMPL.exec(text);
  while (matches) {
    variableName = matches[2];
    if (!(variableName in foundVariables)) {
      let itmKind = CompletionItemKind.Variable;
      if (matches[1].toLowerCase() == "const")
        itmKind = CompletionItemKind.Constant;
      foundVariables[variableName] = true;
      variables.push(createNewCompletionItem(itmKind, variableName));
    }
    matches = PATTERNS.VAR_COMPL.exec(text);
  }

  return variables;
}

function getLocalFunctionCompletions(text: string): CompletionItem[] {
  const functions: CompletionItem[] = [];
  const foundFunctions = {};

  let matches = PATTERNS.FUNC_COMPL.exec(text);
  while (matches) {
    const functionName = matches[3];
    if (!(functionName in foundFunctions)) {
      let itmKind = CompletionItemKind.Function;
      if (matches[1].toLowerCase() == "sub")
        itmKind = CompletionItemKind.Method;
      foundFunctions[functionName] = true;
      functions.push(createNewCompletionItem(itmKind, functionName));
    }
    matches = PATTERNS.FUNC_COMPL.exec(text);
  }

  return functions;
}

function getLocalPropertyCompletions(text: string): CompletionItem[] {
  const vals: CompletionItem[] = [];
  const foundVals = {};

  let matches = PATTERNS.PROP_COMPL.exec(text);
  while (matches) {
    const name = matches[1];
    if (!(name in foundVals)) {
      foundVals[name] = true;
      vals.push(createNewCompletionItem(CompletionItemKind.Property, name));
    }
    matches = PATTERNS.PROP_COMPL.exec(text);
  }

  return vals;
}

function getLocalClassCompletions(text: string): CompletionItem[] {
  const vals: CompletionItem[] = [];
  const foundVals = {};

  let matches = PATTERNS.CLASS.exec(text);
  while (matches) {
    const name = matches[1];
    if (!(name in foundVals)) {
      foundVals[name] = true;
      vals.push(createNewCompletionItem(CompletionItemKind.Class, name));
    }
    matches = PATTERNS.CLASS.exec(text);
  }

  return vals;
}

function provideCompletionItems(document: TextDocument, position: Position) {
  // Gather the functions created by the user
  const text = document.getText();
  let range = document.getWordRangeAtPosition(position);

  if (!range) {
    range = new Range(position, position);
  }

  // Remove completion offerings from commented lines
  const line = document.lineAt(position);
  const firstChar = line.text.charAt(line.firstNonWhitespaceCharacterIndex);
  if (firstChar === "'") {
    return null;
  }

  const VAR = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class))[\t ]+([a-z_0-9]+)\b/i; //fix: should again after var name
  if (VAR.exec(line.text))
    return;

  const variableCompletions = getVariableCompletions(text);
  const functionCompletions = getLocalFunctionCompletions(text);
  const propertyCompletions = getLocalPropertyCompletions(text);
  const classCompletions = getLocalClassCompletions(text);

  return [...definitions, ...variableCompletions, ...functionCompletions, ...propertyCompletions, ...classCompletions];
}

export default languages.registerCompletionItemProvider(
  { scheme: 'file', language: 'vbs' },
  { provideCompletionItems },
);
