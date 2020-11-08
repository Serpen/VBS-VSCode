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

export const customIncludeDirs = workspace.getConfiguration("vbs").get<string[]>("customIncludeDirs");
export const custumIncludePattern = new RegExp(workspace.getConfiguration("vbs").get<string>("custumIncludePattern"), "ig");

export function getImportsWithLocal(doc : TextDocument) : [string, IncludeFile][] {
  const localIncludes = [...Includes];
  const processedMatches = Array<string>();

  let match : RegExpExecArray;
  while ((match = custumIncludePattern.exec(doc.getText())) !== null) {
    if (processedMatches.indexOf(match[1].toLowerCase())) {
      for (const incDir of customIncludeDirs) {
        let incDirResolved = incDir;
        if (incDirResolved === ".")
          incDirResolved = workspace.workspaceFolders[0].uri.fsPath;
        if (incDirResolved === "..")
          incDirResolved = pathns.dirname(workspace.workspaceFolders[0].uri.fsPath);

        const path = pathns.resolve(incDirResolved, match[1]);
        if (fs.existsSync(path) && fs.statSync(path)?.isFile())
          localIncludes.push([
            `Import Statement ${match[1]}`,
            new IncludeFile(path)
          ]);
        else if (fs.existsSync(`${path }.vbs`) && fs.statSync(`${path }.vbs`)?.isFile())
          localIncludes.push([
            `Import Statement ${match[1]}`,
            new IncludeFile(`${path}.vbs`)
          ]);
      }
      processedMatches.push(match[1].toLowerCase());
    }
  }

  return localIncludes;
}
