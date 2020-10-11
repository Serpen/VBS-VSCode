# VBScript Extension for Visual Studio Code
This extension implements basic language features of Visual Basic Script/VBScript/VBS for [Visual Studio Code](https://code.visualstudio.com/).

## Features
- Outline / Gliederung
- Completion / Autovervollständigung

![Outline](assets/docs/Completion-And-Outline.png)
- Goto Definition / Gehe zu Definition
- Run (no debugging)
- Hover / 

![Hover](assets/docs/Hover.png)
- Signatures / Signaturen

![Hover](assets/docs/Signature.png)

- Add extra VBS Source (libraries) files for extra completion
```
{ // settings.json
    "vbs.includes": ["c:\\mylibrary.vbs"]
}
```

## Configuration
| Config                 | Description                          
|------------------------|--------------------------------------|
|vbs.interpreter         | Path to Script Interpreter           |
|vbs.includes            | Array of additional VBS Source Files |
|vbs.showVariableSymbols | Show Variables in Outline            |

## References / Thanks
This project is based on the Visual Basic extension shipped with VS Code and the AutoIt Extension https://github.com/loganch/AutoIt-VSCode.

### Purpose
This project was founded to help developing with VBS in an buisness application (medico Klinische Dokumentation)