import functions from "./functions.json";
import keywords from "./keywords.json";
import constants from "./constants.json";
import colorconstants from "./colorconstants.json";
import operators from "./operators.json"
import props from './props.json'
import types from "./types.json"

import { CompletionItem, CompletionItemKind } from "vscode";

var completions = new Array<CompletionItem>();

for (let entry in functions) {
    const itm = new CompletionItem(entry, CompletionItemKind.Function);
    itm.detail = functions[entry]?.label;
    itm.documentation = functions[entry]?.documentation;
    completions.push(itm);
}

for (let entry in keywords) {
    const itm = new CompletionItem(entry, CompletionItemKind.Keyword);
    itm.detail = entry;
    itm.documentation = keywords[entry]?.documentation;
    completions.push(itm);
}

for (let entry in constants) {
    const itm = new CompletionItem(entry, CompletionItemKind.Constant);
    itm.detail = "Const " + entry;
    itm.documentation = constants[entry]?.documentation;
    completions.push(itm);
}
for (let entry in colorconstants) {
    const itm = new CompletionItem(entry, CompletionItemKind.Color);
    itm.detail = "Const " + entry;
    itm.documentation = colorconstants[entry]?.documentation;
    completions.push(itm);
}

for (let entry in operators) {
    const itm = new CompletionItem(entry, CompletionItemKind.Operator);
    itm.detail = entry;
    itm.documentation = operators[entry]?.documentation;
    completions.push(itm);
}

for (let entry in props) {
    const itm = new CompletionItem(entry, CompletionItemKind.Property);
    itm.detail = "Property " + entry;
    itm.documentation = props[entry]?.documentation;
    completions.push(itm);
}

for (let entry in types) {
    const itm = new CompletionItem(entry, CompletionItemKind.Struct);
    itm.detail = "Type " + entry;
    itm.documentation = types[entry]?.documentation;
    completions.push(itm);
}

export default completions
