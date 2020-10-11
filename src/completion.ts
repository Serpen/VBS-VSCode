import { languages, CompletionItem, CompletionItemKind, TextDocument, Position} from 'vscode';
import definitions from './definitions';
import { GlobalSourceImport, ObjectSourceImport, SourceImports } from './extension';
import * as PATTERNS from './patterns';

function getVariableCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = new Array<string>();

  let matches: RegExpExecArray;
  while ((matches = PATTERNS.VAR.exec(text)) !== null) {
    matches[2].split(",").forEach(match => {
      const name = match.replace(PATTERNS.ARRAYBRACKETS, '').trim();

      if (!(name in foundVals)) {
        let itmKind = CompletionItemKind.Variable;
        if (matches[1].toLowerCase().indexOf("const") > 0)
          itmKind = CompletionItemKind.Constant;

        const ci = new CompletionItem(name, itmKind);
        ci.documentation = matches[3];

        ci.detail = matches[0] + ` [${scope}]`;

        foundVals.push(name);
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
  while ((matches = PATTERNS.FUNCTION.exec(text)) !== null) {
    const functionName = matches[5];

    if (!(functionName in foundVals)) {
      let itmKind = CompletionItemKind.Function;
      if (matches[3].toLowerCase() === "sub")
        itmKind = CompletionItemKind.Method;
      const ci = new CompletionItem(functionName, itmKind);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

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

      foundVals.push(functionName);
      CIs.push(ci);
    }
  }

  return CIs;
}

function getPropertyCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = new Array<string>();

  let matches: RegExpExecArray;
  while ((matches = PATTERNS.PROP.exec(text)) !== null) {
    const name = matches[4];

    if (!(name in foundVals)) {
      const ci = new CompletionItem(name, CompletionItemKind.Property);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = matches[2] + ` [${scope}]`;

      foundVals.push(name);
      CIs.push(ci);
    }
  }

  return CIs;
}

function getClassCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = new Array<string>();

  let matches;
  while ((matches = PATTERNS.CLASS.exec(text)) !== null) {
    const name = matches[3];
    if (!(name in foundVals)) {
      foundVals.push(name);
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
  return [...getVariableCompletions(text, scope), ...getFunctionCompletions(text, scope, parseParams), ...getPropertyCompletions(text, scope), ...getClassCompletions(text, scope)];
}

function provideCompletionItems(document: TextDocument, position: Position): CompletionItem[] {
  // Gather the functions created by the user
  const text = document.getText();
  const codeAtPosition = document.lineAt(position).text.substring(0, position.character);

  // Remove completion offerings from commented lines
  const line = document.lineAt(position);
  if (line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
    return [];

  if (PATTERNS.VAR_COMPLS.test(codeAtPosition))
    return [];

  const retCI: CompletionItem[] = [];

  if (/\s+\.$/.test(codeAtPosition))
    return [];

  let count = 0;
  for (const cp of codeAtPosition)
    if (cp === '"') count++;
  if (count % 2 === 1)
    return [];

  if (/.*\.\w*$/.test(codeAtPosition)) {
    if (/.*\bErr\.\w*$/i.test(codeAtPosition)) {
      const Obj = /Class Err.*?End Class/is.exec(ObjectSourceImport);
      retCI.push(...getFunctionCompletions(Obj[0], "Err"), ...getPropertyCompletions(Obj[0], "Err"));
    } else if (/.*\bWScript\.\w*$/i.test(codeAtPosition)) {
      const Obj = /Class WScript.*?End Class/is.exec(ObjectSourceImport);
      retCI.push(...getFunctionCompletions(Obj[0], "WScript"), ...getPropertyCompletions(Obj[0], "WScript"));
    } else if (/.*fso\.\w*$/i.test(codeAtPosition)) { // dirty!
      const Obj = /Class FileSystemObject.*?End Class/is.exec(ObjectSourceImport);
      retCI.push(...getFunctionCompletions(Obj[0], "FSO"), ...getPropertyCompletions(Obj[0], "FSO"));
    } else {
      retCI.push(...getCompletions(text, "Local"));

      retCI.push(...getFunctionCompletions(ObjectSourceImport, "Object"), ...getPropertyCompletions(ObjectSourceImport, "Object"));

      SourceImports.forEach(ImportDoc => {
        retCI.push(...getCompletions(ImportDoc, "Import"));
      });
    }
  } else {
    retCI.push(...definitions);
    retCI.push(...getCompletions(text, "Local", true));

    retCI.push(...getClassCompletions(ObjectSourceImport, "Global"));

    retCI.push(...getCompletions(GlobalSourceImport, "Global"));

    SourceImports.forEach(ImportDoc => {
      retCI.push(...getCompletions(ImportDoc, "Import"));
    });
  }

  return retCI;
}

export default languages.registerCompletionItemProvider(
  { scheme: 'file', language: 'vbs' },
  { provideCompletionItems }, "."
);
