import * as vs from "vscode";
import * as patterns from "./patterns";

function provideDocumentFormattingEdits(doc: vs.TextDocument, options: vs.FormattingOptions): vs.ProviderResult<vs.TextEdit[]> {
  const textEdits : vs.TextEdit[] = [];

  const FUNCTION = new RegExp(patterns.FUNCTION, "i");
  const CLASS = new RegExp(patterns.CLASS, "i");
  const PROP = new RegExp(patterns.PROP, "i");
  const ENDLINE = new RegExp(patterns.ENDLINE, "i");

  const INDENTCHAR = "    ";

  let indent = 0;
  for (let l = 0; l < doc.lineCount; l++) {
    const line = doc.lineAt(l).text;

    if (line.trim() !== "" && !(/^[ \t]*'/).test(line)) {
      if (ENDLINE.test(line))
        indent--;
      const match = line.match(new RegExp(`\\${INDENTCHAR}{${indent}}(\\w.*)$`));
      if (!match) {
        const r = new vs.Range(l, 0, l, line.length);
        const newText = INDENTCHAR.repeat(indent) + line.trim();
        textEdits.push(new vs.TextEdit(r, newText));
      }

      if (FUNCTION.test(line) || CLASS.test(line) || PROP.test(line))
        indent++;
    }

  }

  return textEdits;
}


export default vs.languages.registerDocumentFormattingEditProvider({ scheme: "file", language: "vbs" }, { provideDocumentFormattingEdits });

/*
 * Author: Demon
 * Date: 2011/12/27
 * Link: http://demon.tw/tool/vbs-beautifier.html
*/

// function VbsBeautifier() {
//   let quoted: string | any[], comments: string | any[], code: string, indents = {},
//     ReservedWord: string, BuiltInFunction: string, BuiltInConstants: string, VersionInfo;

//   ReservedWord = "And As Boolean ByRef Byte ByVal Call Case Class Const Currency Debug Dim Do Double Each Else ElseIf Empty End EndIf Enum Eqv Event Exit Explicit False For Function Get Goto If Imp Implements In Integer Is Let Like Long Loop LSet Me Mod New Next Not Nothing Null On Option Optional Or ParamArray Preserve Private Property Public RaiseEvent ReDim Rem Resume RSet Select Set Shared Single Static Stop Sub Then To True Type TypeOf Until Variant WEnd While With Xor";
//   BuiltInFunction = "Abs Array Asc Atn CBool CByte CCur CDate CDbl CInt CLng CSng CStr Chr Cos CreateObject Date DateAdd DateDiff DatePart DateSerial DateValue Day Escape Eval Exp Filter Fix FormatCurrency FormatDateTime FormatNumber FormatPercent GetLocale GetObject GetRef Hex Hour InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase LTrim Left Len LoadPicture Log Mid Minute Month MonthName MsgBox Now Oct Randomize RGB RTrim Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second SetLocale Sgn Sin Space Split Sqr StrComp StrReverse String Tan Time TimeSerial TimeValue Timer Trim TypeName UBound UCase Unescape VarType Weekday WeekdayName Year";
//   BuiltInConstants = "vbBlack vbRed vbGreen vbYellow vbBlue vbMagenta vbCyan vbWhite vbBinaryCompare vbTextCompare vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday vbUseSystemDayOfWeek vbFirstJan1 vbFirstFourDays vbFirstFullWeek vbGeneralDate vbLongDate vbShortDate vbLongTime vbShortTime vbObjectError vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation vbDefaultButton1 vbDefaultButton2 vbDefaultButton3 vbDefaultButton4 vbApplicationModal vbSystemModal vbOK vbCancel vbAbort vbRetry vbIgnore vbYes vbNo vbCr vbCrLf vbFormFeed vbLf vbNewLine vbNullChar vbNullString vbTab vbVerticalTab vbUseDefault vbTrue vbFalse vbEmpty vbNull vbInteger vbLong vbSingle vbDouble vbCurrency vbDate vbString vbObject vbError vbBoolean vbVariant vbDataObject vbDecimal vbByte vbArray WScript";

//   indents.if = 1;
//   indents.sub = 1;
//   indents.function = 1;
//   indents.property = 1;
//   indents.for = 1;
//   indents.while = 1;
//   indents.do = 1;
//   indents.for = 1;
//   indents.select = 1;
//   indents.with = 1;
//   indents.class = 1;
//   indents.end = -1;
//   indents.next = -1;
//   indents.loop = -1;
//   indents.wend = -1;

//   this.beautify = function(input: any) {
//     code = input;
//     code = code.replace(/\r\n/g, "\n");
//     getQuoted();
//     getComments();
//     getErrorHandling();

//     colonToNewLine();
//     fixSpaces();
//     replaceReservedWord();
//     insertIndent();
//     fixIndent();

//     putErrorHandling();
//     putComments();
//     putQuoted();
//     code = code.replace(/\n/g, "\r\n").replace(/\t/g, "    ");

//     return code;
//   };

//   function getQuoted() {
//     quoted = code.match(/".*?"/g);
//     code = code.replace(/".*?"/g, "%[quoted]%");
//   }

//   function putQuoted() {
//     if (!quoted) return;
//     for (let i = 0, l: number = quoted.length; i < l; i++) {
//       code = code.replace("%[quoted]%", quoted[i]);
//     }
//   }

//   function getComments() {
//     comments = code.match(/'.*/g);
//     code = code.replace(/'.*/g, "%[comment]%");
//   }

//   function putComments() {
//     if (!comments) return;
//     for (let i = 0, l: number = comments.length; i < l; i++) {
//       code = code.replace("%[comment]%", comments[i]);
//     }
//   }

//   function getErrorHandling() {
//     code = code.replace(/on\s+error\s+resume\s+next/gi, "%[resumenext]%");
//     code = code.replace(/on\s+error\s+goto\s+0/gi, "%[gotozero]%");
//   }

//   function putErrorHandling() {
//     code = code.replace(/%\[resumenext\]%/gi, "On Error Resume Next");
//     code = code.replace(/%\[gotozero\]%/gi, "On Error GoTo 0");
//   }

//   function colonToNewLine() {
//     code = code.replace(/:/g, "\n");
//   }

//   function fixSpaces() {
//     code = code.
//       replace(/^[ \t]*(.*?)[ \t]*$/gim, "$1").
//       replace(/[ \t]*(=|<|>|-|\+|&|\*|\/|\^|\\)[ \t]*/gim, " $1 ").
//       replace(/[ \t]*<\s*>[ \t]*/gim, " <> ").
//       replace(/[ \t]*<\s*=[ \t]*/gim, " <= ").
//       replace(/[ \t]*>\s*=[ \t]*/gim, " >= ").
//       replace(/[ \t]*_[ \t]*$/gim, " _").
//       replace(/[ \t]*(Do|Loop)\s*(While|Until)[ \t]*/gim, "$1 $2").
//       replace(/[ \t]*End\s*(Sub|Function|Property|If|With|Select|Class)[ \t]*/gim, "End $1").
//       replace(/[ \t]*Select\s*Case[ \t]*/gim, "Select Case ");
//   }

//   function replaceReservedWord() {
//     let a, i, l, r;
//     a = `${ReservedWord } ${ BuiltInFunction } ${ BuiltInConstants}`;
//     a = a.split(" ");
//     for (i = 0, l = a.length; i < l; i++) {
//       r = new RegExp(`\\b${ a[i] }\\b`, "gi");
//       code = code.replace(r, a[i]);
//     }
//   }

//   function insertIndent() {
//     let a, i, l, line, t = 0, delta = 0;
//     a = code.split("\n");
//     for (i = 0, l = a.length; i < l; i++) {
//       line = a[i];
//       line = cutSomeWords(line);
//       t = delta;
//       delta += countDelta(line);
//       if (t <= delta) {
//         a[i] = str_repeat("\t", t) + a[i];
//       } else {
//         a[i] = str_repeat("\t", delta) + a[i];
//       }
//     }
//     code = a.join("\n");
//   }

//   function cutSomeWords(s: string) {
//     return s.replace(/if.*?then.+/i, "").
//       replace(/(private|public).+?(sub|function|property)/i, "$2");
//   }

//   function countDelta(s: string) {
//     let i, r;
//     for (i in indents) {
//       r = new RegExp(`^[ \\t]*\\b${ i }\\b`, "i");
//       if (s.match(r)) {
//         return indents[i];
//       }
//     }

//     return 0;
//   }

//   function fixIndent() {
//     let a, i, l;
//     a = code.split("\n");
//     for (i = 0, l = a.length; i < l; i++) {
//       if (a[i].match(/^\t*else/i)) {
//         a[i] = a[i].replace("\t", "");
//       }
//     }
//     code = a.join("\n");
//   }

//   function str_repeat(input: string, multiplier: number) {
//     if (multiplier < 0) multiplier = 0;

//     return new Array(multiplier + 1).join(input);
//   }
// }
