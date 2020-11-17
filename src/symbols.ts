/* eslint-disable complexity */
/* eslint-disable max-statements */
import { languages, SymbolKind, DocumentSymbol, Range, workspace, TextDocument } from "vscode";
import * as PATTERNS from "./patterns";

const showVariableSymbols: boolean = workspace.getConfiguration("vbs").get<boolean>("showVariableSymbols");
const showParameterSymbols: boolean = workspace.getConfiguration("vbs").get<boolean>("showParamSymbols");

const FUNCTION = RegExp(PATTERNS.FUNCTION.source, "i");
const CLASS = RegExp(PATTERNS.CLASS.source, "i");
const PROP = RegExp(PATTERNS.PROP.source, "i");

function provideDocumentSymbols(doc: TextDocument): DocumentSymbol[] {
  const result: DocumentSymbol[] = [];

  const varList: string[] = [];

  const Blocks: DocumentSymbol[] = [];

  for (let lineNum = 0; lineNum < doc.lineCount; lineNum++) {
    const line = doc.lineAt(lineNum);

    if (line.isEmptyOrWhitespace || line.text.charAt(line.firstNonWhitespaceCharacterIndex) === "'")
      continue;

    const LineTextwithoutComment = (/^([^'\n\r]*).*$/m).exec(line.text);

    for (const lineText of LineTextwithoutComment[1].split(":")) {
      let name: string;
      let symbol: DocumentSymbol | null;

      let matches: RegExpMatchArray | null = [];

      if ((matches = CLASS.exec(lineText)) !== null) {
        name = matches[3];
        symbol = new DocumentSymbol(name, "", SymbolKind.Class, line.range, line.range);

      } else if ((matches = FUNCTION.exec(lineText)) !== null) {
        name = matches[4];
        let detail = "";
        let symKind = SymbolKind.Function;
        if (matches[3].toLowerCase() === "sub")
          if ((/class_(initialize|terminate)/i).test(name)) {
            symKind = SymbolKind.Constructor;
          } else {
            detail = "Sub";
          }
        else {
          detail = "Function";
        }

        // if params are shown extra, def line shouldn't contain it too
        if (showParameterSymbols)
          name = matches[5];

        symbol = new DocumentSymbol(name, detail, symKind, line.range, line.range);

        if (showParameterSymbols) {
          if (matches[6])
            matches[6].split(",").forEach(param => {
              symbol.children.push(new DocumentSymbol(param.trim(), "Parameter", SymbolKind.Variable, line.range, line.range));
            });
        }

      } else if ((matches = PROP.exec(lineText)) !== null) {
        name = matches[4];
        symbol = new DocumentSymbol(name, matches[3], SymbolKind.Property, line.range, line.range);
        if ((/Default[\t ]*Property[\t ]*Get/i).test(matches[2]))
          symbol.detail = "Default Get";

      } else if (showVariableSymbols) {
        while ((matches = PATTERNS.VAR.exec(lineText)) !== null) {
          const varNames = matches[2].split(",");
          for (const varname of varNames) {
            const vname = varname.replace(PATTERNS.ARRAYBRACKETS, "").trim();
            if (varList.indexOf(vname) === -1 || !(/\bSet\b/i).test(matches[0])) { // match multiple same Dim, but not an additional set to a dim
              varList.push(vname);
              let symKind = SymbolKind.Variable;
              if ((/\bconst\b/i).test(matches[1]))
                symKind = SymbolKind.Constant;
              else if ((/\bSet\b/i).test(matches[0]))
                symKind = SymbolKind.Struct;
              else if ((/\w+[\t ]*\([\t ]*\d*[\t ]*\)/i).test(varname))
                symKind = SymbolKind.Array;
              const r = new Range(line.lineNumber, 0, line.lineNumber, PATTERNS.VAR.lastIndex);
              const variableSymbol = new DocumentSymbol(vname, "", symKind, r, r);
              if (Blocks.length === 0)
                result.push(variableSymbol);
              else
                Blocks[Blocks.length - 1].children.push(variableSymbol);
            }
          }
        }
      }

      if (symbol) {
        if (Blocks.length === 0)
          result.push(symbol);
        else
          Blocks[Blocks.length - 1].children.push(symbol);
        Blocks.push(symbol);
      }

      if ((matches = PATTERNS.ENDLINE.exec(lineText)) !== null)
        Blocks.pop();
    }
  } // next linenum

  return result;
}

export default languages.registerDocumentSymbolProvider(
  { scheme: "file", language: "vbs" },
  { provideDocumentSymbols }
);

