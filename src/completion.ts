import { languages, CompletionItem, CompletionItemKind, TextDocument, Position, CompletionContext } from 'vscode';
import definitions from './definitions';
import { GlobalSourceImport, ObjectSourceImport, SourceImports } from './extension';
import * as PATTERNS from './patterns';

function getVariableCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundVals = {};

  let matches: RegExpMatchArray;
  while ((matches = PATTERNS.VAR.exec(text)) !== null) {
    matches[2].split(",").forEach(match => {
      const name = match.trim();

      if (!(name in foundVals)) {
        let itmKind = CompletionItemKind.Variable;
        if (matches[1].toLowerCase().indexOf("const") > 0)
          itmKind = CompletionItemKind.Constant;

        const ci = new CompletionItem(name, itmKind);
        ci.documentation = matches[3];

        ci.detail = matches[0] + ` [${scope}]`;

        foundVals[name] = true;
        CIs.push(ci);
      }
    });
  }

  return CIs;
}

function getFunctionCompletions(text: string, scope: string): CompletionItem[] {
  const CIs: CompletionItem[] = [];
  const foundFunctions = {};

  let matches: RegExpExecArray;
  while ((matches = PATTERNS.FUNCTION.exec(text)) !== null) {
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

  let matches: RegExpMatchArray;
  while ((matches = PATTERNS.PROP.exec(text)) !== null) {
    const name = matches[4];

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

  let matches;
  while ((matches = PATTERNS.CLASS.exec(text)) !== null) {
    const name = matches[2];
    if (!(name in foundVals)) {
      foundVals[name] = true;
      let ci = new CompletionItem(name, CompletionItemKind.Class);

      if (matches[1]) {
        const summary = PATTERNS.COMMENT_SUMMARY.exec(matches[1]);
        ci.documentation = summary?.[1];
      }

      ci.detail = `[${scope}] ` + name
      CIs.push(ci);
    }
  }

  return CIs;
}

function getCompletions(text: string, scope: string) {
  return [...getVariableCompletions(text, scope), ...getFunctionCompletions(text, scope), ...getPropertyCompletions(text, scope), ...getClassCompletions(text, scope)];

}

function provideCompletionItems(document: TextDocument, position: Position, _token, context: CompletionContext): CompletionItem[] {
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

  let count: number = 0;
  for (let i = 0; i < codeAtPosition.length; i++)
    if (codeAtPosition[i] == '"') count++;
  if (count % 2 == 1)
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
    retCI.push(...getCompletions(text, "Local"));

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
