const FUNCTION         = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z0-9_]+)\((.*)\))/img; 
const FUNCTION_COMMENT = /((?:^[\t ]*'''.*(?:\n|\r\n))+)*[\t ]*((?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z]\w+)\((.*)\)))\s*$/img;

const CLASS = /^[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-z0-9_]+)/img;

const PROP         = /^[\t ]*(?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z0-9_]+)/img;
const PROP_COMMENT = /((?:^[\t ]*'''.*(?:\n|\r\n))+)*[\t ]*((?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(?:Get|Let|Set)[\t ]+([a-z]\w+))/img;


const VAR = /(?:^|:)[\t ]*(Dim|(?:Private |Public )?Const)[\t ]+([a-z0-9_]+)\b/img;
const VAR2 = /(?<!'\s*)(?:^|:)[\t ]*(?:Dim|Const|Private Const|Public Const|Private|Public)[\t ]+(?!Sub|Function|Class|Property)([a-z0-9_]+(?:[\t ]*,[\t ]*[a-z0-9_]+)*)[\t ]*/img;
const VAR_COMMENT = /(?:^|:)[\t ]*((Dim|(?:Private[\t ]*|Public[\t ]*)?Const)[\t ]+([a-z]\w*)(?:\s*=\s*[^'\n\r]+)?)(?:'\s*(.+))?$/img;

const DEF = (word : string) => new RegExp(`((?<!^.*(?:'|Rem)\\s*)(?:Class|Const|Dim|Function|Property [GLS]et|Sub)[\\t ]+${word}[^'\\r\\n]*).*$`, "im")

const COMMENT_SUMMARY = /'''\s*<summary>(.*)<\/summary>/i;
const PARAM_SUMMARY = (input : string, word : string) => new RegExp(`'''\\s*<param name="${word}">(.*)<\\/param>`, "i").exec(input);

const PATTERNS = {
  FUNCTION,
  FUNCTION_COMMENT,
  VAR,
  VAR2,
  VAR_COMMENT,
  CLASS,
  PROP,
  PROP_COMMENT,
  DEF,
  COMMENT_SUMMARY,
  PARAM_SUMMARY
};

export default PATTERNS;