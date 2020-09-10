'use strict'
import { CompletionItem, CompletionItemKind } from "vscode";

var items: CompletionItem[] = [
    {
        label: 'Array',
        documentation: ''
    },
    {
        label: 'Boolean',
        documentation: ''
    },
    {
        label: 'Byte',
        documentation: ''
    },
    {
        label: 'Currency',
        documentation: ''
    },
    {
        label: 'Integer',
        documentation: ''
    },
    {
        label: 'Long',
        documentation: ''
    },
    {
        label: 'Single',
        documentation: ''
    },
    {
        label: 'String',
        documentation: ''
    },
    {
        label: 'Variant',
        documentation: ''
    },
]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Struct
    i.detail = 'Datatype'
    // i.insertText = new SnippetString(i.insertText)
}

export default items;