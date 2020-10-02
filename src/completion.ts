import fs from 'fs';
import { languages, CompletionItem, CompletionItemKind, Range, TextDocument, Position, workspace } from 'vscode';
import definitions from './definitions';
import PATTERNS from './patterns';

function createNewCompletionItem(kind: CompletionItemKind, name: string, strDetail = '') {
  const compItem = new CompletionItem(name, kind);

  if (strDetail != '')
    compItem.detail = strDetail;
  else {
    compItem.detail = "Document " + CompletionItemKind[compItem.kind];
  }
  return compItem;
}

function getVariableCompletions(text: string): CompletionItem[] {
  const variables: CompletionItem[] = [];
  const foundVariables = {};
  let variableName: string;

  let matches = PATTERNS.VAR.exec(text);
  while (matches) {
    variableName = matches[2];
    if (!(variableName in foundVariables)) {
      let itmKind = CompletionItemKind.Variable;
      if (matches[1].toLowerCase().endsWith("const"))
        itmKind = CompletionItemKind.Constant;
      foundVariables[variableName] = true;
      variables.push(createNewCompletionItem(itmKind, variableName));
    }
    matches = PATTERNS.VAR.exec(text);
  }

  return variables;
}

function getFunctionCompletions(text: string): CompletionItem[] {
  const functions: CompletionItem[] = [];
  const foundFunctions = {};

  const FUNCTION = /((?:^[\t ]*'''.*(?:\n|\r\n))+)*[\t ]*((?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z]\w+)\((.*)\)))\s*$/img;
  const SUMMARY = /'''\s*<summary>(.*)<\/summary>/i

  let matches = FUNCTION.exec(text);
  while (matches) {
    const functionName = matches[5];
    console.log("getFunctionCompletions "+ functionName);
    
    if (!(functionName in foundFunctions)) {
      let itmKind = CompletionItemKind.Function;
      if (matches[3].toLowerCase() == "sub")
        itmKind = CompletionItemKind.Method;
      const ci = new CompletionItem(functionName, itmKind);
     
      if (matches[1]) {
        const summary = SUMMARY.exec(matches[1]);
        ci.documentation = summary[1];
      }
      
      ci.detail = "[Global] " + matches[2];

      foundFunctions[functionName] = true;
      functions.push(ci);
    }
    matches = FUNCTION.exec(text);
  }

  return functions;
}

function getPropertyCompletions(text: string): CompletionItem[] {
  const vals: CompletionItem[] = [];
  const foundVals = {};

  let matches = PATTERNS.PROP.exec(text);
  while (matches) {
    const name = matches[2];
    if (!(name in foundVals)) {
      foundVals[name] = true;
      vals.push(createNewCompletionItem(CompletionItemKind.Property, name));
    }
    matches = PATTERNS.PROP.exec(text);
  }

  return vals;
}

function getClassCompletions(text: string): CompletionItem[] {
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

  if (!range)
    range = new Range(position, position);

  // Remove completion offerings from commented lines
  const line = document.lineAt(position);
  if (line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
    return null;

  const VAR = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class|Property [GLT]et))[\t ]+/i; //fix: should again after var name
  if (VAR.test(line.text))
    return;

  const variableCompletions = getVariableCompletions(text);
  const functionCompletions = getFunctionCompletions(text);
  const propertyCompletions = getPropertyCompletions(text);
  const classCompletions = getClassCompletions(text);

  const ExtraDocument: string = workspace.getConfiguration("vbs").get("includes");

  let extracompl: CompletionItem[] = [];
  if (ExtraDocument != '' && fs.statSync(ExtraDocument)) {
    const exttext = fs.readFileSync(ExtraDocument).toString();
    extracompl = [...getFunctionCompletions(exttext), ...getPropertyCompletions(exttext), ...getClassCompletions(exttext)];

    extracompl.forEach(element => {
      if (element.detail == null) {
        element.detail = "Included " + CompletionItemKind[element.kind];
      }
    });
  }

  return [...definitions, ...variableCompletions, ...functionCompletions, ...propertyCompletions, ...classCompletions, ...extracompl];
}

export default languages.registerCompletionItemProvider(
  { scheme: 'file', language: 'vbs' },
  { provideCompletionItems }, " "
);
