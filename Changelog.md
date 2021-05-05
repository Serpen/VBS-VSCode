# VBS-VSCode Changelog

## 1.2.1

### Bugfixes
- customIncludeDirs now support ., .. and ${workspaceFolder} correctly #52
- default customIncludeDirs is now . and ${workspaceFolder}
- custumIncludePattern config (wrong spelling) deprecated, use customIncludePattern
- multiple same include statements, won't show included completion multiple times
- show parameters in outline config
- Support VSCode >= 1.47.0
- patched (irrelevant) Depandapot security warning for y18n


### Known Issues
- Most parts of the extension aren't context sensitiv #18
  A definition/hover etc. only matches the correct name, the scope is ignored


## 1.2.0
### Features
- Use VSCodes launch config, instead of command for running scripts, see #48 for more info
- Generate Diagnostics from a failed Script run
- Show color Information inline
- Config for user definied include function, to provider completion
- Provide Localization for English and German
- Extension needs VSCode ^1.43.0

### Bugfixes
- Code tidy with more eslint
- Show octal and hex numbers, date in syntax
- Multi Variable Syntax coloring
- Better If Snippet Completion

### Known Issues
- Most parts of the extension aren't context sensitiv #18
  A definition/hover etc. only matches the correct name, the scope is ignored


## 1.1.0
### Features
- More support for comment based help and where it is shown #14, #25, #34
- Parameter values are shown in hover, completion,  goto definition and optional in symbols #12
- Some more language documentation #21 and snippets
- Support VSCode >= 1.33.0

### Bugfixes
- lots of small RegExp fixes/regressions, the should be better oberserved with automated tests via mocha and appveyor
- Autosave before run is now waiting to be saved #31
- internal code fixes, for stricter code

### Known Issues
- Most parts of the extension aren't context sensitiv #18
  A definition/hover etc. only matches the correct name, the scope is ignored

## 1.0
- First public stable Release