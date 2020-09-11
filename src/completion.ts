import { languages, CompletionItem, CompletionItemKind, Range, TextDocument, Position } from 'vscode';
import definitions from './definitions';
import PATTERNS from './patterns';
import UTIL from './util';

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

/**
 * Creates an array of completion items for AutoIt variables from the document.
 * @param {String} text Content of the document
 * @returns {Array<Object>} Array of CompletionItem objects
 */
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

/**
 * Creates an array of CompletionItems for Functions declared in the document
 * @param {String} text Content of the document
 * @returns {Array<Object>} Array of CompletionItem objects
 */
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

const provideCompletionItems = (document: TextDocument, position: Position) => {
  // Gather the functions created by the user

  const text = document.getText();
  let range = document.getWordRangeAtPosition(position);

  if (!range) {
    range = new Range(position, position);
  }

  // Remove completion offerings from commented lines
  const line = document.lineAt(position.line);
  const firstChar = line.text.charAt(line.firstNonWhitespaceCharacterIndex);
  if (firstChar === ';') {
    return null;
  }

  const variableCompletions = getVariableCompletions(text);
  const functionCompletions = getLocalFunctionCompletions(text);
  const propertyCompletions = getLocalPropertyCompletions(text);
  const classCompletions = getLocalClassCompletions(text);

  const localCompletions = [...variableCompletions, ...functionCompletions, ...propertyCompletions, ...classCompletions];

  return [...definitions, ...localCompletions];
};

module.exports = languages.registerCompletionItemProvider(
  UTIL.VBS_MODE,
  { provideCompletionItems },
  '.',
);
