export const FUNCTION = /((?:^[\t ]*'+.*(?:\n|\r\n))+)*[\t ]*((?:Public[\t ]+|Private[\t ]+)?(Function|Sub)[\t ]+(([a-z]\w*)[\t ]*(?:\((.*)\))?))\s*/img;

export const CLASS = /((?:^[\t ]*'+.*(?:\n|\r\n))+)*[\t ]*(?:Public[\t ]+|Private[\t ]+)?Class[\t ]+([a-z]\w*)/img;

export const PROP = /((?:^[\t ]*'+.*(?:\n|\r\n))+)*[\t ]*((?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z]\w*))/img;

export const VAR = /(?<!'\s*)(?:^|:)[\t ]*(Dim|Set|Const|Private[\t ]+Const|Public[\t ]+Const|Private|Public)[\t ]+(?!Sub|Function|Class|Property)([a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?(?:[\t ]*,[\t ]*[a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?)*)[\t ]*.*(?:$|:)/img;
export const VAR_COMPLS = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class|Property [GLT]et))[\t ]+/i; //fix: should again after var name #22

export function DEF(input: string, word: string) {
  return new RegExp(
    `^[^'\\n\\r]*(\\b(?:Class|Const|Dim|Function|Property[\\t ][GLS]et|Sub)\\b[\\t ]+[\\w,\\t ]*\\b${word}\\b(?:[\\t ]*\(.*\))?)(?<!.*\\bRem\\b.*)`
    , "im").exec(input);
}

export const COMMENT_SUMMARY = /(?:'''\s*<summary>|'\s*)([^<\n\r]*)(?:<\/summary>)?/i;
export const PARAM_SUMMARY = (input: string, word: string) => new RegExp(`'''\\s*<param name=["']${word}["']>(.*)<\\/param>`, "i").exec(input);

export const ENDLINE = (/(?:^|:)[\t ]*End\s+(Sub|Class|Function|Property)/i);

export const ARRAYBRACKETS = /\(\s*\d*\s*\)/;

