# VBS-VSCode Changelog

## 1.1.0
### Features
- More support for comment based help and where it is shown #14, #25, #34
- Parameter values are shown in hover, completion,  goto definition and optional in symbols #12
- Some more language documentation #21 and snippets


### Bugfixes
- lots of small RegExp fixes/regressions, the should be better oberserved with automated tests via mocha and appveyor
- Autosave before run is now waiting to be saved #31
- internal code fixes, for stricter code

### Known Issues
- Most parts of the extension aren't context sensitiv #18
  A definition/hover etc. only matches the correct name, the scope is ignored

## 1.0
- First public stable Release