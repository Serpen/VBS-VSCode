const FUNCTION = /^[\t ]{0,}(?:Public\s+|Private\s+)?(Function|Sub)\s+(.+)\(/i;
const VAR = /^[\t ]{0,}(Dim|Const)\s+(\w+)/i;

const INCLUDE = /^#include\s"(.+)"/gm;
const FUNCTION_SIG = /(?=\S)(?!;~\s)(?:Function|Sub)\s+((\w+)\((.+)?\))/gi;
const FUNCTION_SIG_LOCAL = /^[\t ]{0,}(?:Function|Sub)\s+((\w+)\((.+)?\))/gmi;
    
const Function = (word : string) => new RegExp(`\\b(Function|Sub|Dim|Const|Class)\\s+${word}\\b`, "i");

const PATTERNS = {
  FUNCTION,
  VAR,
  FUNCTION_SIG,
  FUNCTION_SIG_LOCAL,
  INCLUDE,
  Function
};

export default PATTERNS;