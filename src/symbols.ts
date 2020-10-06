import { languages, SymbolKind, DocumentSymbol, Range, workspace } from 'vscode';
import * as PATTERNS from './patterns';

export default languages.registerDocumentSymbolProvider({ scheme: 'file', language: 'vbs' }, {
  provideDocumentSymbols(doc) {
    const result: DocumentSymbol[] = [];

    const FUNCTION = RegExp(PATTERNS.FUNCTION.source, 'i');
    const CLASS = RegExp(PATTERNS.CLASS.source, 'i');
    const PROP = RegExp(PATTERNS.PROP.source, 'i');

    const varList: String[] = [];

    const showVariableSymbols: boolean = workspace.getConfiguration("vbs").get<boolean>("showVariableSymbols")!;

    let Blocks: DocumentSymbol[] = [];
    let BlockEnds: String[] = [];

    // Get the number of lines in the document to loop through
    const lineCount = Math.min(doc.lineCount, 10000);
    for (let lineNum = 0; lineNum < lineCount; lineNum++) {
      const line = doc.lineAt(lineNum);


      if (line.isEmptyOrWhitespace || line.text.charAt(line.firstNonWhitespaceCharacterIndex) == "'")
        continue;

      const LineTextwithoutComment = (/^([^'\n\r]*).*$/m).exec(line.text);

      if (LineTextwithoutComment?.[1]) {
        LineTextwithoutComment[1].split(":").forEach(lineText => {
          let name: string;
          let symbol: DocumentSymbol | null;

          let matches: RegExpMatchArray | null = [];

          if ((matches = CLASS.exec(lineText)) !== null) {
            name = matches[2];
            symbol = new DocumentSymbol(name, '', SymbolKind.Class, line.range, line.range);
            BlockEnds.push("class")

          } else if ((matches = FUNCTION.exec(lineText)) !== null) {
            name = matches[4];
            let detail: string = "";
            let symKind = SymbolKind.Function;
            if (matches[3].toLowerCase() === "sub")
              if (name.toLowerCase() == "class_initialize()" || name.toLowerCase() == "class_terminate()") {
                symKind = SymbolKind.Constructor;
                BlockEnds.push("sub");
              } else {
                detail = "Sub";
                BlockEnds.push(detail.toLowerCase());
              }
            else {
              detail = "Function";
              BlockEnds.push(detail.toLowerCase());
            }
            symbol = new DocumentSymbol(name, detail, symKind, line.range, line.range);
            
          } else if ((matches = PROP.exec(lineText)) !== null) {
            name = matches[4];
            symbol = new DocumentSymbol(name, matches[3], SymbolKind.Property, line.range, line.range);
            BlockEnds.push("property");

          } else if (showVariableSymbols) {
            while ((matches = PATTERNS.VAR.exec(lineText)) !== null) {
              const varNames = matches[2].split(',');
              for (let i = 0; i < varNames.length; i++) {
                let name = varNames[i].trim();
                if (varList.indexOf(name) == -1 || !(/\bSet\b/i.test(matches[0]))) { // match multiple same Dim, but not an additional set to a dim
                  varList.push(name);
                  let symKind = SymbolKind.Variable;
                  if (/\bconst\b/i.test(matches[1]))
                    symKind = SymbolKind.Constant;
                  else if (/\bSet\b/i.test(matches[0]))
                    symKind = SymbolKind.Struct;
                  else if (/\w+[\t ]*\([\t ]*\d*[\t ]*\)/i.test(name))
                    symKind = SymbolKind.Array;
                  let r = new Range(line.lineNumber, 0, line.lineNumber, PATTERNS.VAR.lastIndex);
                  const variableSymbol = new DocumentSymbol(name, '', symKind, r, r);
                  if (Blocks.length == 0)
                    result.push(variableSymbol);
                  else
                    Blocks[Blocks.length - 1].children.push(variableSymbol);
                }
              }
            }
          }

          if (symbol!) {
            if (Blocks.length == 0)
              result.push(symbol!);
            else
              Blocks[Blocks.length - 1].children.push(symbol!);
            Blocks.push(symbol!);
          }

          if ((matches = PATTERNS.ENDLINE.exec(lineText)) !== null)
            if (BlockEnds[BlockEnds.length - 1] == matches[1].toLowerCase()) {
              Blocks.pop();
              BlockEnds.pop();
            } else {
              Blocks.pop();
              console.log("symbol wrong ending (awaiting closing for " + BlockEnds.pop()?.toString() + " got " + matches[1].toLowerCase() + ") in " + doc.uri + " " + line.lineNumber)
            }

        });
      }
    } // next linenum

    if (BlockEnds.length > 0)
      console.log(BlockEnds);
    if (Blocks.length > 0)
      console.log(Blocks);

    return result;
  },
});

