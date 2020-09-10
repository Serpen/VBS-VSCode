'use strict'
import { CompletionItem, CompletionItemKind } from "vscode";

var items: CompletionItem[] = [
    {
        label: 'ByRef',
        documentation: ''
    },
    {
        label: 'ByVal',
        documentation: ''
    },
    {
        label: 'Call',
        documentation: ''
    },
    {
        label: 'Case',
        documentation: ''
    },
    {
        label: 'Class',
        documentation: ''
    },
    {
        label: 'Const',
        documentation: ''
    },
    {
        label: 'Dim',
        documentation: ''
    },
    {
        label: 'Each',
        documentation: ''
    },
    {
        label: 'Else',
        documentation: ''
    },
    {
        label: 'End',
        documentation: ''
    },
    {
        label: 'Enum',
        documentation: ''
    },
    {
        label: 'Event',
        documentation: ''
    },
    {
        label: 'Exit',
        documentation: ''
    },
    {
        label: 'Function',
        documentation: ''
    },
    {
        label: 'Get',
        documentation: ''
    },
    {
        label: 'GoTo',
        documentation: ''
    },
    {
        label: 'Implements',
        documentation: ''
    },
    {
        label: 'Let',
        documentation: ''
    },
    {
        label: 'Loop',
        documentation: ''
    },
    {
        label: 'New',
        documentation: ''
    },
    {
        label: 'Next',
        documentation: ''
    },
    {
        label: 'Nothing',
        documentation: ''
    },
    {
        label: 'Null',
        documentation: ''
    },
    {
        label: 'Option',
        documentation: ''
    },
    {
        label: 'Optional',
        documentation: ''
    },
    {
        label: 'ParamArray',
        documentation: ''
    },
    {
        label: 'Preserve',
        documentation: ''
    },
    {
        label: 'Private',
        documentation: ''
    },
    {
        label: 'Public',
        documentation: ''
    },
    {
        label: 'RaiseEvent',
        documentation: ''
    },
    {
        label: 'ReDim',
        documentation: ''
    },
    {
        label: 'Rem',
        documentation: ''
    },
    {
        label: 'Resume',
        documentation: ''
    },
    {
        label: 'Select',
        documentation: ''
    },
    {
        label: 'Set',
        documentation: ''
    },
    {
        label: 'Shared',
        documentation: ''
    },
    {
        label: 'Static',
        documentation: ''
    },
    {
        label: 'Stop',
        documentation: ''
    },
    {
        label: 'Sub',
        documentation: ''
    },
    {
        label: 'Then',
        documentation: ''
    },
    {
        label: 'Type',
        documentation: ''
    },
    {
        label: 'TypeOf',
        documentation: ''
    },
    {
        label: 'Until',
        documentation: ''
    },
    {
        label: 'Wend',
        documentation: ''
    },
    {
        label: 'While',
        documentation: ''
    },
    {
        label: 'With',
        documentation: ''
    },
]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Keyword
    i.detail = 'Keyword'
    // i.insertText = new SnippetString(i.insertText)
}

export default items;