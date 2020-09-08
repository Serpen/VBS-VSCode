import fs from 'fs';
import path from 'path';
import {TextDocument } from 'vscode';

const VBS_MODE = { language: 'vbs', scheme: 'file' };

function arraysMatch(arr1, arr2) {
  if (arr1.length === arr2.length && arr1.some(v => arr2.indexOf(v) <= 0)) {
    return true;
  }
  return false;
}

const getIncludeText = (filePath : string) => {
  return fs.readFileSync(filePath).toString();
};

const getIncludePath = (fileOrPath : string, document : TextDocument) : string => {
  let includePath = '';

  if (fileOrPath.charAt(1) === ':') {
    includePath = fileOrPath;
  } else {
    let docDir = path.dirname(document.fileName);

    docDir +=
      (fileOrPath.charAt(0) === '\\' || fileOrPath.charAt(0) === '/' ? '' : '\\') + fileOrPath;
    includePath = path.normalize(docDir);
  }

  includePath = includePath.charAt(0).toUpperCase() + includePath.slice(1);

  return includePath;
};



const UTIL = {
  VBS_MODE,
  getIncludeText,
  getIncludePath,
  arraysMatch
};

export default UTIL;