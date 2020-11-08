# VBScript Extension für Visual Studio Code
Diese Erweiterung bietet Sprachunterstützung für Visual Basic Script/VBScript/VBS für [Visual Studio Code](https://code.visualstudio.com/).

[![Version](https://vsmarketplacebadge.apphb.com/version/serpen.vbsvscode.svg)](https://marketplace.visualstudio.com/items?itemName=serpen.vbsvscode)
[![Build status](https://ci.appveyor.com/api/projects/status/0i0hrbt657y8geef?svg=true)](https://ci.appveyor.com/project/Serpen/vbs-vscode)

<p align="center">
  <a href="./README.md">English</a> | 
  <span>Deutsch</span>
</p>

## Features
- Gliederung
- Autovervollständigung

![Outline](assets/docs/Completion-And-Outline.png)
- Gehe zu Definition
- Ausführen (kein Debugging)
- Hover

![Hover](assets/docs/Hover.png)
- Signaturen

![Hover](assets/docs/Signature.png)

- Farbinformationen anzeigen

![Hover](assets/docs/ColorProvider.png)

- Zusätzliche VBS Funktions Bibliotheken als VBS Dateien einbinden
```
{ // settings.json
    "vbs.includes": ["c:\\mylibrary.vbs"]
}
```

## Mitarbeit
Du kannst dieses Projekt unterstützen, indem du die Quelldateien forkst und einen Pull Request/PR mit deinen Veränderungen erzeugst oder eine Issue mit deinem Problem/deiner Idee erzeugst.
- Vervollständigung der VBS Sprachdokumentation #21
- Übersetzung in weitere Sprachen
- ...


## Referenzen / Danksagung
Diese Erweiterung basiert auf der Visual Basic Erweitung die mit VS Code ausgeliefert wird und dem Design der AutoIt Erweiterung von loganch.

## Zweck
Die Erweiterung wurde entwickelt um die VBS Entwicklung innerhalb einer Geschäftssoftware (medico Klinische Dokumentation) zu unterstützen.
