# VBScript Extension for Visual Studio Code
This extension implements basic language features of Visual Basic Script/VBScript/VBS for [Visual Studio Code](https://code.visualstudio.com/).

[![Version](https://vsmarketplacebadge.apphb.com/version/serpen.vbsvscode.svg)](https://marketplace.visualstudio.com/items?itemName=serpen.vbsvscode)
[![Build status](https://ci.appveyor.com/api/projects/status/0i0hrbt657y8geef?svg=true)](https://ci.appveyor.com/project/Serpen/vbs-vscode)

<p align="center">
  <span>English</span> | 
  <a href="./README.de.md">Deutsch</a>
</p>

## Features
- Outline
- Completion

![Outline](assets/docs/Completion-And-Outline.png)
- Goto Definition
- Run (no debugging)
- Hover 

![Hover](assets/docs/Hover.png)

- Signatures
![Hover](assets/docs/Signature.png)

- Color Information

![Hover](assets/docs/ColorProvider.png)

- Add extra VBS Source (libraries) files for extra completion
```
{ // settings.json
    "vbs.includes": ["c:\\mylibrary.vbs"]
}
```

## Contribute
You can support this project through PR with your changes or simply add an issue with your idea/bug.
- Complete Language Source Documentation #21
- Translate
- ...

## References / Thanks
This extension is based on the Visual Basic extension shipped with VS Code and the design from AutoIt Extension from loganch.

## Purpose
This extension was founded to help developing with VBS in an buisness application (medico Klinische Dokumentation)