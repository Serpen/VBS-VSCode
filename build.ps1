New-Item $env:temp\Serpen-vbs-vscode.x.x.x -ItemType Directory
Copy-Item -Path .\dist, *.vbs, .\package.json, .\LICENSE, .\README.md, .\snippets, .\syntaxes -Recurse -Destination $env:temp\Serpen-vbs-vscode.x.x.x\
Compress-Archive $env:temp\Serpen-vbs-vscode.x.x.x -DestinationPath $PSScriptRoot\..\Serpen-vbs-vscode.x.x.x.zip -CompressionLevel Optimal -Force
Del $env:temp\Serpen-vbs-vscode.x.x.x\ -Recurse -Force