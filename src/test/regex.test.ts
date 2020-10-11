import { strictEqual } from 'assert';
import * as PATTERNS from '../patterns';

describe('Function matching', () => {
  const matchresult = PATTERNS.FUNCTION.exec("' Comment\nfunction myFunction(param1)");
  it('parse Function comment', () => {
    strictEqual(matchresult[1].trim(), "' Comment")
  });
  it('parse Function definition', () => {
    strictEqual(matchresult[2], "function myFunction(param1)")
  });
  it('parse function type', () => {
    strictEqual(matchresult[3], "function")
  });
  it('parse function def2', () => {
    strictEqual(matchresult[4], "myFunction(param1)")
  });
  it('parse function name', () => {
    strictEqual(matchresult[5], "myFunction")
  });
  it('parse function param', () => {
    strictEqual(matchresult[6], "param1")
  });
});


describe('Class matching', () => {
  const matchresult = PATTERNS.CLASS.exec("' Comment\nClass myClass");
  it('parse Class comment', () => {
    strictEqual(matchresult[1].trim(), "' Comment")
  });
  it('parse Class def', () => {
    strictEqual(matchresult[2], "Class myClass")
  });
  it('parse Class name', () => {
    strictEqual(matchresult[3], "myClass")
  });
});


describe('Property matching', () => {
  const matchresult = PATTERNS.PROP.exec("' Comment\nPublic Property Get myProperty");
  it('parse Property comment', () => {
    strictEqual(matchresult[1].trim(), "' Comment")
  });
  it('parse Property definition', () => {
    strictEqual(matchresult[2], "Public Property Get myProperty")
  });
  it('parse Property type', () => {
    strictEqual(matchresult[3], "Get")
  });
  it('parse Property name', () => {
    strictEqual(matchresult[4], "myProperty")
  });
  it('parse function params', () => {
    strictEqual(matchresult[5], undefined)
  });
});

describe('Variable matching', () => {
  const matchresult = PATTERNS.VAR.exec("Dim varname ' Comment");
  it('parse Property comment', () => {
    strictEqual(matchresult[1], "Dim")
  });
  it('parse Property definition', () => {
    strictEqual(matchresult[2], "varname")
  });
});
