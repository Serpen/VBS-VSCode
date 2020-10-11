import { strictEqual } from 'assert';
import * as PATTERNS from '../patterns';

describe('Function matching', () => {
  const testString = "' Comment\nfunction myFunction(param1)";
  const match = PATTERNS.FUNCTION.exec(testString);

  
  it('parse Function comment', () => {
    strictEqual(match[1].trim(), "' Comment")
  });
  it('parse Function definition', () => {
    strictEqual(match[2], "function myFunction(param1)")
  });
  it('parse function type', () => {
    strictEqual(match[3], "function")
  });
  it('parse function sig def', () => {
    strictEqual(match[4], "myFunction(param1)")
  });
  it('parse function name', () => {
    strictEqual(match[5], "myFunction")
  });
  it('parse function param', () => {
    strictEqual(match[6], "param1")
  });
});


describe('Class matching', () => {
  const match = PATTERNS.CLASS.exec("' Comment\nClass myClass");

  it('parse Class comment', () => {
    strictEqual(match[1].trim(), "' Comment")
  });
  it('parse Class def', () => {
    strictEqual(match[2], "Class myClass")
  });
  it('parse Class name', () => {
    strictEqual(match[3], "myClass")
  });
});


describe('Property matching', () => {
  const match = PATTERNS.PROP.exec("' Comment\nPublic Property Get myProperty");

  it('parse Property comment', () => {
    strictEqual(match[1].trim(), "' Comment")
  });
  it('parse Property definition', () => {
    strictEqual(match[2], "Public Property Get myProperty")
  });
  it('parse Property type', () => {
    strictEqual(match[3], "Get")
  });
  it('parse Property name', () => {
    strictEqual(match[4], "myProperty")
  });
  it('parse function params', () => {
    strictEqual(match[5], undefined)
  });
});

describe('Variable matching', () => {
  const match = PATTERNS.VAR.exec("Dim varname ' Comment");

  it('parse Property comment', () => {
    strictEqual(match[1], "Dim")
  });
  it('parse Property definition', () => {
    strictEqual(match[2], "varname")
  });
});
