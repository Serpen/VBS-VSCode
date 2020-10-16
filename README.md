# VBScript Extension for Visual Studio Code
This extension implements basic language features of Visual Basic Script/VBScript/VBS for [Visual Studio Code](https://code.visualstudio.com/).

[![Version](https://vsmarketplacebadge.apphb.com/version/serpen.vbsvscode.svg)](https://marketplace.visualstudio.com/items?itemName=serpen.vbsvscode)
[![Build status](https://ci.appveyor.com/api/projects/status/0i0hrbt657y8geef?svg=true)](https://ci.appveyor.com/project/Serpen/vbs-vscode)


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

## References / Thanks
This project is based on the Visual Basic extension shipped with VS Code and the AutoIt Extension from loganch.

### Purpose
This project was founded to help developing with VBS in an buisness application (medico Klinische Dokumentation)