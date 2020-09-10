
import functions from "./func_def.json";
import keywords from "./keywords.json";
import constants from "./constants.json";
import colorconstants from "./colorconstants.json";
import operators from "./operators.json"
import props from './props.json'
import types from "./types.json"

import { CompletionItem, CompletionItemKind } from "vscode";

// Add the function icon and detail to each entry

var completions = new Array<CompletionItem>();

for (let i in functions) {
    const itm = new CompletionItem(i, CompletionItemKind.Function);
    itm.detail = functions[i]?.label;
    itm.documentation = functions[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}

for (let i in keywords) {
    const itm = new CompletionItem(i, CompletionItemKind.Keyword);
    itm.detail = i;
    itm.documentation = keywords[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}

for (let i in constants) {
    const itm = new CompletionItem(i, CompletionItemKind.Constant);
    itm.detail = i;
    itm.documentation = constants[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}
for (let i in colorconstants) {
    const itm = new CompletionItem(i, CompletionItemKind.Color);
    itm.detail = i;
    itm.documentation = colorconstants[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}

for (let i in operators) {
    const itm = new CompletionItem(i, CompletionItemKind.Operator);
    itm.detail = i;
    itm.documentation = operators[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}

for (let i in props) {
    const itm = new CompletionItem(i, CompletionItemKind.Property);
    itm.detail = i;
    itm.documentation = props[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}

for (let i in types) {
    const itm = new CompletionItem(i, CompletionItemKind.Property);
    itm.detail = i;
    itm.documentation = types[i]?.documentation;
    // i.insertText = new SnippetString(i.insertText)
    completions.push(itm);
}

export default completions
