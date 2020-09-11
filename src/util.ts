import { DocumentSelector } from 'vscode';

const VBS_MODE: DocumentSelector = { scheme: 'file', language: 'vbs' };

function arraysMatch(arr1: [], arr2: []) {
  if (arr1.length === arr2.length && arr1.some(v => arr2.indexOf(v) <= 0)) {
    return true;
  }
  return false;
}

const UTIL = {
  VBS_MODE,
  arraysMatch
};

export default UTIL;