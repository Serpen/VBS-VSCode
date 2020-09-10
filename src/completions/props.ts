'use strict'
import { CompletionItem, CompletionItemKind } from "vscode";

var items: CompletionItem[] = [
    {
        label: 'Date',
        documentation: ''
    },
    {
        label: 'Debug',
        documentation: ''
    },
    {
        label: 'Err',
        documentation: 'Class'
    },
    {
        label: 'GetLocale',
        documentation: ''
    },
    {
        label: 'GetUILanguage',
        documentation: ''
    },
    {
        label: 'Now',
        documentation: ''
    },
    {
        label: 'ScriptEngine',
        documentation: ''
    },
    {
        label: 'ScriptEngineBuildVersion',
        documentation: ''
    },
    {
        label: 'ScriptEngineMajorVersion',
        documentation: ''
    },
    {
        label: 'ScriptEngineMinorVersion',
        documentation: ''
    },
    {
        label: 'Time',
        documentation: ''
    },
    {
        label: 'Timer',
        documentation: ''
    },
    {
        label: 'WScript',
        documentation: ''
    },
    
]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Property
    i.detail = 'Property'
    // i.insertText = new SnippetString(i.insertText)
}

export default items;