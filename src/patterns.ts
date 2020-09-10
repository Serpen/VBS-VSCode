const FUNCTION = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-zA-Z0-9_]+)\((.*)\))/img; 
const FUNC_DEF = (word : string) => new RegExp(`(?:Function|Sub)\\s${word}\\(`, "i");
const FUNC_INC = /^(?=\S)(?!;~\s)(?:Function|Sub)\s+([a-z_0-9]+)\s*\(/gmi;
const FUNC_COMPL = /\b(Function|Sub)\s+([a-z_0-9]+)\s*\(/gi;
const FUNCTION_SIG = /^[\t ]{0,}(?:(?:Public|Private)\s+)?(?:Function|Sub)\s+(([a-z_0-9]+)\((.+)?\))/img;
const Function = (word : string) => new RegExp(`\\b(Function|Sub|Dim|Const|Class)\\s+${word}\\b`, "i");

const CLASS = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-zA-Z0-9_]+)/img;

const PROP = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Property[\t ]+(?:Get|Let|Set)[\t ]+([a-z_0-9]+)/img;

const VAR_DEF = (word : string) => new RegExp(`(?:Dim|Const)\\s${word}\\s?=?`, 'i');
const VAR = /(?:^|:)[\t ]*(Dim|Const)[\t ]+([a-z_0-9]+)\b/img;

const LIBRARY_INCLUDE = /^#include\s+<([\w.]+\.vbs)>/gm;
const INCLUDE = /^#include\s"(.+)"/gm;

const PATTERNS = {
  FUNCTION,
  VAR,
  VAR_COMPL : VAR,
  PROP_COMPL : PROP,
  CLASS,
  PROP,
  FUNCTION_SIG : FUNCTION,
  INCLUDE,
  LIBRARY_INCLUDE,
  FUNC_INC : FUNCTION,
  FUNC_COMPL : FUNCTION,
  VAR_DEF,
  FUNC_DEF,
  Function
};

export default PATTERNS;