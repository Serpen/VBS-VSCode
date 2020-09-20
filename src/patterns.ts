const FUNCTION = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z0-9_]+)\((.*)\))/img; 

const CLASS = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-z0-9_]+)/img;

const PROP = /^[\t ]*(?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z0-9_]+)/img;

const VAR = /(?:^|:)[\t ]*(Dim|(?:Private |Public )?Const)[\t ]+([a-z0-9_]+)\b/img;

const DEF = (word : string) => new RegExp(`(Class|Const|Dim|Function|Property [GLS]et|Sub)[\\t ]+${word}`, "i")

const PATTERNS = {
  FUNCTION,
  VAR,
  CLASS,
  PROP,
  DEF
};

export default PATTERNS;