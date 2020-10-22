import { languages, CompletionItem, CompletionItemKind, TextDocument, Position } from "vscode";
import definitions from "./definitions";
import { Includes } from "./extension";
import * as PATTERNS from "./patterns";

function getVariableCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = []; // results
  const foundVals = new Array<string>(); // list to prevent doubles

  let matches: RegExpExecArray;
  while (matches = PATTERNS.VAR.exec(text)) {
    matches[2].split(",").forEach(match => {
      const name = match.replace(PATTERNS.ARRAYBRACKETS, "").trim();

      if (foundVals.indexOf(name.toLowerCase()) === -1) {
        foundVals.push(name.toLowerCase());

        let itmKind = CompletionItemKind.Variable;
        if (/\bconst\b/i.test(matches[1]))
          itmKind = CompletionItemKind.Constant;

        const ci = new CompletionItem(name, itmKind);
        ci.documentation = matches[3];

        ci.detail = matches[0] + ` [${scope}]`;

        CIs.push(ci);
      }
    });
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

      let itmKind = CompletionItemKind.Function;
      if (matches[3].toLowerCase() === "sub")
        itmKind = CompletionItemKind.Method;
      const ci = new CompletionItem(name, itmKind);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      // currently only parse in local document, but for all functions, since there is no context
      if (parseParams)
        matches[6]?.split(",").forEach(param => {
          const paramCI = new CompletionItem(param.trim(), CompletionItemKind.Variable);
          if (matches[1]) {
            const paramComment = PATTERNS.PARAM_SUMMARY(matches[1], param.trim());
            if (paramComment)
              paramCI.documentation = paramComment[1];
          }

          CIs.push(paramCI);
        });

      ci.detail = matches[2] + ` [${scope}]`;

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

    if ((foundVals.indexOf(name.toLowerCase()) === -1)) {
      foundVals.push(name.toLowerCase());

      const ci = new CompletionItem(name, CompletionItemKind.Property);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = matches[2] + ` [${scope}]`;

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

      ci.detail = name + ` [${scope}]`;
      CIs.push(ci);
    }
  }

  return CIs;
}

function getCompletions(text: string, scope: string, parseParams = false) {
  return [...getVariableCompletions(text, scope),
    ...getFunctionCompletions(text, scope, parseParams),
    ...getPropertyCompletions(text, scope),
    ...getClassCompletions(text, scope)];
}

function provideCompletionItems(doc: TextDocument, position: Position): CompletionItem[] {
  const codeAtPosition = doc.lineAt(position).text.substring(0, position.character);

  // Remove completion offerings from commented lines
  const line = doc.lineAt(position);
  if (line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
    return [];

  // no Completion during writing a definition
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

  const ObjectSourceImportName = "ObjectDefs";
  const ObjectSourceImport = Includes.get(ObjectSourceImportName);

  // if dot is typed than show only members
  if (/.*\.\w*$/.test(codeAtPosition)) {
    if (/.*\bErr\.\w*$/i.test(codeAtPosition)) {
      const Obj = /Class Err.*?End Class/is.exec(ObjectSourceImport.Content);
      retCI.push(...getFunctionCompletions(Obj[0], "Err"), ...getPropertyCompletions(Obj[0], "Err"));
    } else if (/.*\bWScript\.\w*$/i.test(codeAtPosition)) {
      const Obj = /Class WScript.*?End Class/is.exec(ObjectSourceImport.Content);
      retCI.push(...getFunctionCompletions(Obj[0], "WScript"), ...getPropertyCompletions(Obj[0], "WScript"));
    } else if (/.*fso\.\w*$/i.test(codeAtPosition)) { // dirty!
      const Obj = /Class FileSystemObject.*?End Class/is.exec(ObjectSourceImport.Content);
      retCI.push(...getFunctionCompletions(Obj[0], "FSO"), ...getPropertyCompletions(Obj[0], "FSO"));
    } else {
      retCI.push(...getCompletions(text, "Local"));

      retCI.push(...getFunctionCompletions(ObjectSourceImport.Content, ObjectSourceImportName),
        ...getPropertyCompletions(ObjectSourceImport.Content, ObjectSourceImportName));

      for (const imp of Includes)
        if (imp[0].startsWith("Import"))
          retCI.push(...getCompletions(imp[1].Content, imp[0]));

    }
  } else { // show global members
    retCI.push(...definitions);
    retCI.push(...getCompletions(text, "Local", true));

    retCI.push(...getClassCompletions(ObjectSourceImport.Content, ObjectSourceImportName));

    for (const item of Includes)
      if (item[0].startsWith("Import") || item[0] === "Global")
        retCI.push(...getCompletions(item[1].Content, item[0]));

  }

  return retCI;
}

export default languages.registerCompletionItemProvider(
  { scheme: "file", language: "vbs" },
  { provideCompletionItems }, "."
);
