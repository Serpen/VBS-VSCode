const FUNCTION = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z0-9_]+)\((.*)\))/img; 

const CLASS = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-z0-9_]+)/img;

const PROP = /^[\t ]*(?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z0-9_]+)/img;

const VAR = /(?:^|:)[\t ]*(Dim|(?:Private |Public )?Const)[\t ]+([a-z0-9_]+)\b/img;
const VAR2 = /(?<!'\s*)(?:^|:)[\t ]*(?:Dim|Const|Private Const|Public Const|Private|Public)[\t ]+(?!Sub|Function|Class|Property)([a-z0-9_]+(?:[\t ]*,[\t ]*[a-z0-9_]+)*)[\t ]*/img;

const DEF = (word : string) => new RegExp(`(?<!.*'.*)(?:Class|Const|Dim|Function|Property [GLS]et|Sub)[\\t ]+${word}`, "i")

const PATTERNS = {
  FUNCTION,
  VAR,
  VAR2,
  CLASS,
  PROP,
  DEF
};

export default PATTERNS;