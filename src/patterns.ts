const FUNCTION = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z0-9_]+)\((.*)\))/img; 
const FUNC_DEF = (word : string) => new RegExp(`(?:Function|Sub)\\s${word}\\(`, "i");

const Function = (word : string) => new RegExp(`\\b(Function|Sub|Dim|Const|Class)\\s+${word}\\b`, "i");

const CLASS = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-z0-9_]+)/img;
const CLASS_DEF = (word : string) => new RegExp(`Class[\\t ]+${word}`, "i");

const PROP = /^[\t ]*(?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z0-9_]+)/img;
const PROP_DEF = (word : string) => new RegExp(`(?:Public[\\t ]+(?:Default[\\t ]+)?|Private[\\t ]+)?Property[\\t ]+[GLS]et[\\t ]+${word}`,"i");

const VAR_DEF = (word : string) => new RegExp(`(?:Dim|Const)\\s${word}\\s?=?`, 'i');
const VAR = /(?:^|:)[\t ]*(Dim|(?:Private |Public )?Const)[\t ]+([a-z0-9_]+)\b/img;


const PATTERNS = {
  FUNCTION,
  VAR,
  CLASS,
  PROP,
  VAR_DEF,
  FUNC_DEF,
  PROP_DEF,
  CLASS_DEF,
  Function
};

export default PATTERNS;