import { TextDocument, Uri, workspace } from "vscode";
import * as pathns from "path";
import * as fs from "fs";

export class IncludeFile {
  constructor(path: string) {
    let path2 = path;
    if (!pathns.isAbsolute(path2))
      path2 = pathns.join(workspace.workspaceFolders[0].uri.fsPath, path2);

    this.Uri = Uri.file(path2);

    if (fs.existsSync(path2) && fs.statSync(path2).isFile())
      this.Content = fs.readFileSync(path2).toString();
  }

  Content = "";

  Uri: Uri;
}

export const Includes = new Map<string, IncludeFile>();

export let customIncludeDirs : string[];
export let customIncludePattern : RegExp;

export function reloadImportDocuments() : void {
  const custumIncludePatternStr = workspace.getConfiguration("vbs").get<string>("custumIncludePattern");
  const customIncludePatternStr = workspace.getConfiguration("vbs").get<string>("customIncludePattern");

  customIncludeDirs = workspace.getConfiguration("vbs").get<string[]>("customIncludeDirs");

  if (custumIncludePatternStr !== customIncludePatternStr && custumIncludePatternStr !== "")
    customIncludePattern = new RegExp(custumIncludePatternStr, "ig");
  else
    customIncludePattern = new RegExp(customIncludePatternStr, "ig");

  const SourceImportFiles = workspace.getConfiguration("vbs").get<string[]>("includes");
  for (const key of Includes.keys()) {
    if (key.startsWith("Include"))
      Includes.delete(key);
  }
  SourceImportFiles?.forEach((file) => {
    Includes.set(`Include ${pathns.basename(file)}`, new IncludeFile(file));
  });
}

export function getImportsWithLocal(doc : TextDocument) : [string, IncludeFile][] {
  const localIncludes = [...Includes];
  const processedMatches = Array<string>();

  let match : RegExpExecArray;
  while ((match = customIncludePattern.exec(doc.getText())) !== null) {
    if (processedMatches.indexOf(match[1].toLowerCase()) === -1) {
      for (const incDir of customIncludeDirs) {
        let incDirResolved = incDir;
        if (incDirResolved === ".")
          incDirResolved = pathns.dirname(doc.uri.fsPath);
        else if (incDirResolved === "..")
          incDirResolved = pathns.dirname(pathns.dirname(doc.uri.fsPath));
        // eslint-disable-next-line no-template-curly-in-string
        else if (incDirResolved === "${workspaceFolder}")
          if (workspace.workspaceFolders)
            incDirResolved = workspace.workspaceFolders[0].uri.fsPath;

        const path = pathns.resolve(incDirResolved, match[1]);
        if (fs.existsSync(path) && fs.statSync(path)?.isFile())
          localIncludes.push([
            `Include Statement ${match[1]}`,
            new IncludeFile(path)
          ]);
      }
      processedMatches.push(match[1].toLowerCase());
    }
  }

  return localIncludes;
}
