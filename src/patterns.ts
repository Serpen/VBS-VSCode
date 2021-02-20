/* eslint-disable max-len */
/**
 * Matches a Function, 1st = Comment, 2nd = Definition, 3rd = Function/Sub, 4th = Signature def, 5th = Name, 6th = params
 */
export const FUNCTION = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:(?:Public|Private)[\t ]+)?(Function|Sub)[\t ]+((\[?[a-z]\w*\]?)[\t ]*(?:\((.*)\))?))/img;

/**
 * Matches a Class, 1st = Comment, 2nd = definition, 3rd = Name
 */
export const CLASS = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:(?:Public|Private)[\t ]+)?Class[\t ]+(\[?[a-z]\w*\]?))/img;

/**
 * Matches a Property, 1st = Comment, 2nd = Definition, 3rd = Get/Let/Set, 4th = Name, 5th = params
 */
export const PROP = /((?:^[\t ]*'+.*$(?:\r\n|\n))*)^[\t ]*((?:Public[\t ]+(?:Default[\t ]+)?|Private[\t ]+)?Property[\t ]+(Get|Let|Set)[\t ]+(\[?[a-z]\w*\]?))(?:\((.*)\))?/img;

/**
 * Matches a Variable Declaration, 1st = Type, 2nd = Name (cs)
 */
export const VAR = /(?<!'\s*)(?:^|:)[\t ]*(Dim|Set|Const|Private[\t ]+Const|Public[\t ]+Const|Private|Public)[\t ]+(?!Sub|Function|Class|Property)([a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?(?:[\t ]*,[\t ]*[a-z0-9_]+(?:[\t ]*\([\t ]*\d*[\t ]*\))?)*)[\t ]*.*(?:$|:)/img;

export const VAR_COMPLS = /^[\t ]*(Dim|Const|((Private|Public)[\t ]+)?(Function|Sub|Class|Property [GLT]et))[\t ]+\w+[^:]*$/i; // fix: should again after var name #22

/**
 * Matches a Definition, 1st = Comment, 2nd = Definition, 3rd = Name
 */
export function DEF(input: string, word: string): RegExpExecArray {
  return new RegExp(
    `((?:^[\\t ]*'.*$(?:\\r\\n|\\n))*)^[^'\\n\\r]*^[\\t ]*((?:(?:(?:(?:Private[\\t ]+|Public[\\t ]+)?(?:Class|Function|Sub|Property[\\t ][GLS]et)))[\\t ]+)(\\b${word}\\b).*)$`
    , "im"
  ).exec(input);
}

export function DEFVAR(input: string, word: string): RegExpExecArray {
  return new RegExp(
    `((?:^[\\t ]*'.*$(?:\\r\\n|\\n))*)^[^'\\n\\r]*^[\\t ]*((?:(?:Const|Dim|(?:Private|Private)(?![\\t ]+(?:Sub|Function)))[\\t ]+)[\\w\\t ,]*(\\b${word}\\b).*)$`
    , "im"
  ).exec(input);
}

export const COMMENT_SUMMARY = /(?:'''\s*<summary>|'\s*)([^<\n\r]*)(?:<\/summary>)?/i;

export function PARAM_SUMMARY(input: string, word: string): RegExpExecArray {
  return new RegExp(`'''\\s*<param name=["']${word}["']>(.*)<\\/param>`, "i").exec(input);
}

export const ENDLINE = (/(?:^|:)[\t ]*End\s+(Sub|Class|Function|Property)/i);

export const ARRAYBRACKETS = /\(\s*\d*\s*\)/;

export const COLOR = /\b(vbBlack|vbBlue|vbCyan|vbGreen|vbMagenta|vbRed|vbWhite|vbYellow)\b|\b(RGB[\t ]*\([\t ]*(&h[0-9a-f]+|\d+)[\t ]*,[\t ]*(&h[0-9a-f]+|\d+)[\t ]*,[\t ]*(&h[0-9a-f]+|\d+)[\t ]*\))|(&h[0-9a-f]{6}\b)/ig;

