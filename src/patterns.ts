/**
 * Matches a Function, 1st = Comment, 2nd = Definition, 3rd = Function/Sub, 4th = ?, 5th = Name, 6th = params
 */
export const FUNCTION = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:(?:Public|Private)[\t ]+)?(Function|Sub)[\t ]+(([a-z]\w*)[\t ]*(?:\((.*)\))?))/img;

/**
 * Matches a Class, 1st = Comment, 2nd = Name
 */
export const CLASS    = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:(?:Public|Private)[\t ]+)?Class[\t ]+([a-z]\w*))/img;

/**
 * Matches a Property, 1st = Comment, 2nd = Defintion, 3rd = Get/Let/Set, 4th = Name, 5th = params
 */
export const PROP     = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+([a-z]\w*))(?:\((.*)\))?/img;

export const VAR = /(?<!'\s*)(?:^|:)[\t ]*(Dim|Set|Const|Private[\t ]+Const|Public[\t ]+Const|Private|Public)[\t ]+(?!Sub|Function|Class|Property)([a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?(?:[\t ]*,[\t ]*[a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?)*)[\t ]*.*(?:$|:)/img;
export const VAR_COMPLS = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class|Property [GLT]et))[\t ]+\w+[^:]*$/i; //fix: should again after var name #22

export function DEF(input: string, word: string) {
  return new RegExp(
    `((?:^[\\t ]*'.*$(?:\r\n|\n))*)^[^'\\n\\r]*(\\b(?:Class|Const|Dim|Function|Property[\\t ][GLS]et|Sub)\\b[\\t ]+[\\w,\\t ]*\\b${word}\\b(?:[\\t ]*\(.*\))?)(?<!.*\\bRem\\b.*)`
    , "im").exec(input);
}

export const COMMENT_SUMMARY = /(?:'''\s*<summary>|'\s*)([^<\n\r]*)(?:<\/summary>)?/i;

export function PARAM_SUMMARY(input: string, word: string) {
  return new RegExp(`'''\\s*<param name=["']${word}["']>(.*)<\\/param>`, "i").exec(input);
}

export const ENDLINE = (/(?:^|:)[\t ]*End\s+(Sub|Class|Function|Property)/i);

export const ARRAYBRACKETS = /\(\s*\d*\s*\)/;

