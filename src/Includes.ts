import { TextDocument, Uri, workspace } from "vscode";
import * as pathns from "path";
import * as fs from "fs";

export class IncludeFile {
  constructor(path: string) {
    if (!pathns.isAbsolute(path))
      path = pathns.join(workspace.workspaceFolders[0].uri.fsPath, path);

    this.Uri = Uri.file(path);

    if (fs.existsSync(path) && fs.statSync(path).isFile())
      this.Content = fs.readFileSync(path).toString();
  }
  Content = "";
  Uri: Uri;
}

export const Includes = new Map<string, IncludeFile>();

export const customIncludeDirs = workspace.getConfiguration("vbs").get<string[]>("customIncludeDirs");
export const custumIncludePattern = new RegExp(workspace.getConfiguration("vbs").get<string>("custumIncludePattern"), "ig");

export function GetLocalImports(doc : TextDocument) : [string, IncludeFile][] {
  const localIncludes = [...Includes];
  const incDirs = customIncludeDirs;
    
  const CUSTOM_INC = custumIncludePattern;
  let match : RegExpExecArray;
  while ((match = CUSTOM_INC.exec(doc.getText())) !== null) {
    for (const incDir of incDirs) {
      let incDirResolved = incDir;
      if (incDirResolved === ".") incDirResolved = workspace.workspaceFolders[0].uri.fsPath;
      if (incDirResolved === "..") incDirResolved = pathns.join(workspace.workspaceFolders[0].uri.fsPath, "..");

      const path = pathns.resolve(incDirResolved, match[1]);
      if (fs.existsSync(path) && fs.statSync(path)?.isFile()) 
        localIncludes.push(["Import Statement " + match[1], new IncludeFile(path)]);
    }
  }
  return localIncludes;
}
