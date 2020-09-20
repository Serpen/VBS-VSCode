const FUNCTION = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-zA-Z0-9_]+)\((.*)\))/img; 
const FUNC_DEF = (word : string) => new RegExp(`(?:Function|Sub)\\s${word}\\(`, "i");
const Function = (word : string) => new RegExp(`\\b(Function|Sub|Dim|Const|Class)\\s+${word}\\b`, "i");

const CLASS = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-zA-Z0-9_]+)/img;

const PROP = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z_0-9]+)/img;

const VAR_DEF = (word : string) => new RegExp(`(?:Dim|Const)\\s${word}\\s?=?`, 'i');
const VAR = /(?:^|:)[\t ]*(Dim|Const)[\t ]+([a-z_0-9]+)\b/img;


const PATTERNS = {
  FUNCTION,
  VAR,
  CLASS,
  PROP,
  VAR_DEF,
  FUNC_DEF,
  Function
};

export default PATTERNS;