const FUNCTION = /^[\t ]{0,}(?:Public\s+|Private\s+)?(Function|Sub)\s+(.+)\(/i;
const FUNC_DEF = (word : string) => new RegExp(`(?:Function|Sub)\\s${word}\\(`, "i");
const FUNC_INC = /^(?=\S)(?!;~\s)(?:Function|Sub)\s+(\w+)\s*\(/gm;
const FUNC_COMPL = /\b(Function|Sub)\s+(\w*)\s*\(/gi;
const FUNCTION_SIG = /^[\t ]{0,}(?:(?:Public|Private)\s+)?(?:Function|Sub)\s+((\w+)\((.+)?\))/gmi;
const Function = (word : string) => new RegExp(`\\b(Function|Sub|Dim|Const|Class)\\s+${word}\\b`, "i");
  
const VAR_DEF = (word : string) => new RegExp(`(?:Dim|Const)\\s${word}\\s?=?`, 'i');
const VAR = /\b(Dim|Const)\s+(\w+)/gi;

const LIBRARY_INCLUDE = /^#include\s+<([\w.]+\.vbs)>/gm;
const INCLUDE = /^#include\s"(.+)"/gm;

const PATTERNS = {
  FUNCTION,
  VAR,
  FUNCTION_SIG,
  INCLUDE,
  LIBRARY_INCLUDE,
  FUNC_INC,
  FUNC_COMPL,
  VAR_DEF,
  FUNC_DEF,
  Function
};

export default PATTERNS;