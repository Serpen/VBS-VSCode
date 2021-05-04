import { languages, CompletionItem, CompletionItemKind, TextDocument, Position } from "vscode";
import definitions from "./definitions";
import { Includes, getImportsWithLocal } from "./Includes";
import * as PATTERNS from "./patterns";

const ObjectSourceImportName = "ObjectDefs";

function getVariableCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = []; // results
  const foundVals = new Array<string>(); // list to prevent doubles

  let matches: RegExpExecArray;
  while (matches = PATTERNS.VAR.exec(text)) {
    for (const match of matches[2].split(",")) {
      const name = match.replace(PATTERNS.ARRAYBRACKETS, "").trim();
      if (foundVals.indexOf(name.toLowerCase()) === -1) {
        foundVals.push(name.toLowerCase());

        let itmKind = CompletionItemKind.Variable;

        if ((/\bconst\b/i).test(matches[1]))
          itmKind = CompletionItemKind.Constant;

        const ci = new CompletionItem(name, itmKind);
        ci.documentation = matches[3];

        if (new RegExp(PATTERNS.COLOR, "i").test(name)) {
          ci.kind = CompletionItemKind.Color;
          ci.filterText = `ColorConstants.${name}`;
          ci.insertText = name;
        }

        ci.detail = `${matches[0]} [${scope}]`;

        CIs.push(ci);
      }
    }
  }

  return CIs;
}

function getFunctionCompletions(text: string, scope: string, parseParams = false): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = new Array<string>();

  let matches: RegExpExecArray;
  while (matches = PATTERNS.FUNCTION.exec(text)) {
    const name = matches[5];

    if (foundVals.indexOf(name.toLowerCase()) === -1) {
      foundVals.push(name.toLowerCase());

      const ci = new CompletionItem(name, CompletionItemKind.Function);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      // currently only parse in local document, but for all functions, since there is no context
      if (parseParams && matches[6])
        for (const param of matches[6].split(",")) {
          const paramCI = new CompletionItem(param.trim(), CompletionItemKind.Variable);
          if (matches[1]) {
            const paramComment = PATTERNS.PARAM_SUMMARY(matches[1], param.trim());
            if (paramComment)
              paramCI.documentation = paramComment[1];
          }

          CIs.push(paramCI);
        }

      ci.detail = `${matches[2]} [${scope}]`;

      CIs.push(ci);
    }
  }

  return CIs;
}

function getPropertyCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = new Array<string>();

  let matches: RegExpExecArray;
  while (matches = PATTERNS.PROP.exec(text)) {
    const name = matches[4];

    if (foundVals.indexOf(name.toLowerCase()) === -1) {
      foundVals.push(name.toLowerCase());

      const ci = new CompletionItem(name, CompletionItemKind.Property);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = `${matches[2]} [${scope}]`;

      CIs.push(ci);
    }
  }

  return CIs;
}

function getClassCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = new Array<string>();

  let matches: RegExpExecArray;
  while (matches = PATTERNS.CLASS.exec(text)) {
    const name = matches[3];
    if (foundVals.indexOf(name.toLowerCase()) === -1) {
      foundVals.push(name.toLowerCase());
      const ci = new CompletionItem(name, CompletionItemKind.Class);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = `${name} [${scope}]`;
      CIs.push(ci);
    }
  }

  return CIs;
}

function getCompletions(text: string, scope: string, parseParams = false) {
  return [
    ...getVariableCompletions(text, scope),
    ...getFunctionCompletions(text, scope, parseParams),
    ...getPropertyCompletions(text, scope),
    ...getClassCompletions(text, scope)
  ];
}

function getObjectMembersCode(line: string, code: string, toAddObj : CompletionItem[]): boolean {
  const matches = { "Err": "Err", "WScript": "WScript", "Debug": "Debug", "fso": "FileSystemObject" };
  for (const cls of ["Err", "WScript", "Debug", "fso"]) {
    const lineClsReg = new RegExp(`.*\\b${cls}\\.\\w*`, "i");
    const codeClsReg = new RegExp(`Class[\t ]+${matches[cls]}.+?End[\t ]+Class`, "is");

    if (lineClsReg.test(line)) {
      const classDef = codeClsReg.exec(code);
      toAddObj.push(...getFunctionCompletions(classDef[0], cls), ...getPropertyCompletions(classDef[0], cls));

      return true;
    }
  }

  return false;
}

function provideCompletionItems(doc: TextDocument, position: Position): CompletionItem[] {
  const codeAtPosition = doc.lineAt(position).text.substring(0, position.character);

  // Remove completion offerings from commented lines
  const line = doc.lineAt(position);
  if (line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
    return [];

  // no Completion during writing a definition, still buggy
  if (PATTERNS.VAR_COMPLS.test(codeAtPosition))
    return [];

  // no completion within open string
  let quoteCount = 0;
  for (const char of codeAtPosition)
    if (char === '"') quoteCount++;
  if (quoteCount % 2 === 1)
    return [];

  const text = doc.getText();
  const retCI: CompletionItem[] = [];

  const ObjectSourceImport = Includes.get(ObjectSourceImportName);

  const localIncludes = getImportsWithLocal(doc);

  // if dot is typed than show only members
  if ((/.*\.\w*$/).test(codeAtPosition)) {
    // eslint-disable-next-line no-empty
    if (getObjectMembersCode(codeAtPosition, ObjectSourceImport.Content, retCI)) {} else {
      retCI.push(...getCompletions(text, "Local"));

      retCI.push(
        ...getFunctionCompletions(ObjectSourceImport.Content, ObjectSourceImportName),
        ...getPropertyCompletions(ObjectSourceImport.Content, ObjectSourceImportName)
      );

      for (const imp of localIncludes)
        if (imp[0].startsWith("Include"))
          retCI.push(...getCompletions(imp[1].Content, imp[0]));

    }
  } else { // show global members
    retCI.push(...definitions);
    retCI.push(...getCompletions(text, "Local", true));

    retCI.push(...getClassCompletions(ObjectSourceImport.Content, ObjectSourceImportName));

    for (const item of localIncludes)
      if (item[0].startsWith("Include") || item[0] === "Global")
        retCI.push(...getCompletions(item[1].Content, item[0]));
  }

  return retCI;
}

export default languages.registerCompletionItemProvider(
  { scheme: "file", language: "vbs" },
  { provideCompletionItems }, "."
);
