'use strict'
import { CompletionItem, CompletionItemKind } from "vscode";

var items: CompletionItem[] = [
    {
        label: 'And',
        documentation: '',
    },
    {
        label: 'Like',
        documentation: ''
    },
    {
        label: 'Mod',
        documentation: ''
    },
    {
        label: 'Not',
        documentation: ''
    },
    {
        label: 'Or',
        documentation: ''
    },
    {
        label: 'Xor',
        documentation: ''
    },
]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Operator
    i.detail = 'Operator'
    // i.insertText = new SnippetString(i.insertText)
}

export default items;