import { notStrictEqual, strictEqual } from "assert";
import * as PATTERNS from "../patterns";

describe("Function matching", () => {
  const testString = "' Comment\nfunction myFunction(param1)";
  const match = PATTERNS.FUNCTION.exec(testString);


  it("parse Function comment", () => {
    strictEqual(match[1].trim(), "' Comment");
  });
  it("parse Function definition", () => {
    strictEqual(match[2], "function myFunction(param1)");
  });
  it("parse function type", () => {
    strictEqual(match[3], "function");
  });
  it("parse function sig def", () => {
    strictEqual(match[4], "myFunction(param1)");
  });
  it("parse function name", () => {
    strictEqual(match[5], "myFunction");
  });
  it("parse function param", () => {
    strictEqual(match[6], "param1");
  });
});


describe("Class matching", () => {
  const match = PATTERNS.CLASS.exec("' Comment\nClass myClass");

  it("parse Class comment", () => {
    strictEqual(match[1].trim(), "' Comment");
  });
  it("parse Class def", () => {
    strictEqual(match[2], "Class myClass");
  });
  it("parse Class name", () => {
    strictEqual(match[3], "myClass");
  });
});


describe("Property matching", () => {
  const match = PATTERNS.PROP.exec("' Comment\nPublic Property Get myProperty");

  it("parse Property comment", () => {
    strictEqual(match[1].trim(), "' Comment");
  });
  it("parse Property definition", () => {
    strictEqual(match[2], "Public Property Get myProperty");
  });
  it("parse Property type", () => {
    strictEqual(match[3], "Get");
  });
  it("parse Property name", () => {
    strictEqual(match[4], "myProperty");
  });
  it("parse function params", () => {
    strictEqual(match[5], undefined);
  });
});

describe("Variable matching", () => {
  const match = PATTERNS.VAR.exec("Dim varname ' Comment");

  it("parse Property comment", () => {
    strictEqual(match[1], "Dim");
  });
  it("parse Property definition", () => {
    strictEqual(match[2], "varname");
  });
});

describe("DEF Match", () => {
  it("multiple varnames", () => {
    const match = PATTERNS.DEFVAR("Dim var1, var2", "var2");
    notStrictEqual(match, null);
  });

  it("var with comment", () => {
    const match = PATTERNS.DEFVAR("Dim varname ' Comment", "varname");
    notStrictEqual(match, null);
  });

  it("only part of name", () => {
    const match = PATTERNS.DEFVAR("Dim varname ' Comment", "varnam");
    strictEqual(match, null);
  });

  // is done within function not regex
  // it('Collon multi Dim, 2nd', () => {
  //   const match = PATTERNS.DEFVAR("Dim var1 : dim var2", "var2");
  //   notStrictEqual(match, null, match[0]);
  // });

  it("Collon multi Dim after comment, 2nd", () => {
    const match = PATTERNS.DEFVAR("Dim var1 :' dim Var2", "var2");
    strictEqual(match, null);
  });

  it("Collon multi Dim after comment, 2nd", () => {
    const match = PATTERNS.DEFVAR("Dim var1 ': dim Var2", "var2");
    strictEqual(match, null);
  });

  it("function with comment, no param", () => {
    const match = PATTERNS.DEF("' comment\nfunction myFunc()", "myFunc");
    notStrictEqual(match, null);
  });

  it("function with comment, no brackets", () => {
    const match = PATTERNS.DEF("' comment\nfunction myFunc", "myFunc");
    notStrictEqual(match, null);
  });

  it("public function with params and full doc", () => {
    // eslint-disable-next-line max-len
    const match = PATTERNS.DEF("' <summary>sth</summary><param name=\"param1\">p1</param>\npublic function myFunc(param1, params2) ' comment", "myFunc");
    notStrictEqual(match, null);
  });

});
