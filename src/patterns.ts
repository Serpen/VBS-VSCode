const FUNCTION = /^[\t ]{0,}(?:Function|Sub)\s(.+)\(/;
const INCLUDE = /^#include\s"(.+)"/gm;


const PATTERNS = {
  FUNCTION,
  INCLUDE,
};

export default PATTERNS;