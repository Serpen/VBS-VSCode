import { languages, CompletionItem, CompletionItemKind, workspace, Range } from 'vscode';
import fs from 'fs';
import path from 'path';
import completions from './completions';
import DEFAULT_UDFS from './constants';
import UTILS from './util';
import PATTERNS from './patterns';

let currentIncludeFiles = [];
let includes = [];


const createNewCompletionItem = (kind, name, strDetail = 'Document Function') => {
  const compItem = new CompletionItem(name, kind);

  compItem.detail = kind === CompletionItemKind.Variable ? 'Variable' : strDetail;

  return compItem;
};

/**
 * Checks a filename with the include paths for a valid path
 * @param {string} file - the filename to append to the paths
 * @returns {(string|boolean)} Full path if found to exist or false
 */
const findFilepath = file => {
  const { includePaths } = workspace.getConfiguration('vbs');

  let newPath;
  const pathFound = includePaths.some(iPath => {
    newPath = path.normalize(`${iPath}\\`) + file;
    if (fs.existsSync(newPath)) {
      return true;
    }
    return false;
  });

  if (pathFound && newPath) {
    return newPath;
  }
  return false;
};

/**
 * Returns an array of AutoIt functions found within a VSCode TextDocument
 * @param {string} fileName
 * @param {vscode.TextDocument} document
 * @returns {Array} Array of functions in file
 */
function getIncludeData(fileName, document) {
  const functions = [];
  const filePath = UTILS.getIncludePath(fileName, document);

  let pattern = null;
  const fileData = UTILS.getIncludeText(filePath);

  pattern = PATTERNS.FUNC_INC.exec(fileData);
  do {
    if (pattern) functions.push(pattern[1]);
    pattern = PATTERNS.FUNC_INC.exec(fileData);
  } while (pattern !== null);

  return functions;
}

/**
 * Generates function completions from files included through library paths
 * @param {Array<String>} libraryIncludes Array containing filenames of library includes
 * @param {Object<TextDocument>} doc Originating text document
 * @returns {CompletionItem[]} Array of completionItem objects
 */
const getLibraryFunctions = (libraryIncludes, doc) => {
  const items = [];
  libraryIncludes.forEach(file => {
    const fullPath = findFilepath(file);
    if (fullPath)
      getIncludeData(fullPath, doc).forEach(newFunc => {
        items.push(
          createNewCompletionItem(CompletionItemKind.Function, newFunc, `Function from ${file}`),
        );
      });
  });

  return items;
};

/**
 * Collects the filenames of library includes, filtering out
 * ones that are already default AutoIt UDFs
 * @param {string} docText The contents of the document
 * @returns {Array<string>} Array of library includes
 */
const getLibraryIncludes = docText => {
  const items = [];
  let pattern = PATTERNS.LIBRARY_INCLUDE.exec(docText);
  while (pattern) {
    const filename = pattern[1].replace('.vbs', '');
    if (DEFAULT_UDFS.indexOf(filename) === -1) {
      items.push(pattern[1]);
    }

    pattern = PATTERNS.LIBRARY_INCLUDE.exec(docText);
  }

  return items;
};

/**
 * Creates an array of completion items for AutoIt variables from the document.
 * @param {String} text Content of the document
 * @param {String} firstChar The first character of the text considered for a completion
 * @returns {Array<Object>} Array of CompletionItem objects
 */
const getVariableCompletions = (text : string, firstChar : string) => {
  const variables = [];
  const foundVariables = {};
  let variableName : string;

  let matches = PATTERNS.VAR_COMPL.exec(text);
  while (matches) {
    variableName = matches[2];
    if (!(variableName in foundVariables)) {
      foundVariables[variableName] = true;
      variables.push(createNewCompletionItem(CompletionItemKind.Variable, variableName[2]));
    }
    matches = PATTERNS.VAR_COMPL.exec(text);
  }

  return variables;
};

/**
 * Creates an array of CompletionItems for Functions declared in the document
 * @param {String} text Content of the document
 * @returns {Array<Object>} Array of CompletionItem objects
 */
const getLocalFunctionCompletions = (text : string) => {
  const functions = [];
  const foundFunctions = {};

  let pattern = PATTERNS.FUNC_COMPL.exec(text);
  while (pattern) {
    const { 1: functionName } = pattern;
    if (!(functionName in foundFunctions)) {
      foundFunctions[functionName] = true;
      functions.push(createNewCompletionItem(CompletionItemKind.Function, functionName));
    }
    pattern = PATTERNS.FUNC_COMPL.exec(text);
  }

  return functions;
};

const provideCompletionItems = (document, position) => {
  // Gather the functions created by the user

  const text = document.getText();
  let range = document.getWordRangeAtPosition(position);
  const prefix = range ? document.getText(range)[0] : '';
  const includesCheck = [];

  if (!range) {
    range = new Range(position, position);
  }

  // Remove completion offerings from commented lines
  const line = document.lineAt(position.line);
  const firstChar = line.text.charAt(line.firstNonWhitespaceCharacterIndex);
  if (firstChar === ';') {
    return null;
  }

  const variableCompletions = getVariableCompletions(text, prefix);
  const functionCompletions = getLocalFunctionCompletions(text);

  const localCompletions = [...variableCompletions, ...functionCompletions];

  // collect the includes of the document
  let pattern = PATTERNS.INCLUDE.exec(text);
  while (pattern) {
    includesCheck.push(pattern[1]);
    pattern = PATTERNS.INCLUDE.exec(text);
  }

  // Redo the include collecting if the includes are different
  if (!UTILS.arraysMatch(includesCheck, currentIncludeFiles)) {
    includes = [];
    let includeFunctions = [];
    includesCheck.forEach(include => {
      includeFunctions = getIncludeData(include, document);
      if (includeFunctions) {
        includeFunctions.forEach(newFunc => {
          includes.push(
            createNewCompletionItem(
              CompletionItemKind.Function,
              newFunc,
              `Function from ${include}`,
            ),
          );
        });
      }
    });

    currentIncludeFiles = includesCheck;
  }

  const libraryIncludes = getLibraryIncludes(text);
  const libraryCompletions = getLibraryFunctions(libraryIncludes, document);

  return [...completions, ...localCompletions, ...includes, ...libraryCompletions];
};

module.exports = languages.registerCompletionItemProvider(
  { language: 'vbs', scheme: 'file' },
  { provideCompletionItems },
  '.',
);
