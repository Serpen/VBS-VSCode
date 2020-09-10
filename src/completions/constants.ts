import { CompletionItem, CompletionItemKind } from "vscode";

var colorItems: CompletionItem[] = [
    {
        label: 'vbBlack',
        documentation: ''
    },
    {
        label: 'vbBlue',
        documentation: ''
    },
    {
        label: 'vbCyan',
        documentation: ''
    },
    {
        label: 'vbGreen',
        documentation: ''
    },
    {
        label: 'vbMagenta',
        documentation: ''
    },
    {
        label: 'vbRed',
        documentation: ''
    },
    {
        label: 'vbWhite',
        documentation: ''
    },
    {
        label: 'vbYellow',
        documentation: ''
    }
]

var items: CompletionItem[] = [
    {
        label: 'False',
        documentation: ''
    },
    {
        label: 'True',
        documentation: ''
    },
    {
        label: 'vbAbort',
        documentation: ''
    },
    {
        label: 'vbAbortRetryIgnore',
        documentation: ''
    },
    {
        label: 'vbApplicationModal',
        documentation: ''
    },
    {
        label: 'vbArray',
        documentation: ''
    },
    {
        label: 'vbBinaryCompare',
        documentation: ''
    },
    {
        label: 'vbBoolean',
        documentation: ''
    },
    {
        label: 'vbByte',
        documentation: ''
    },
    {
        label: 'vbCancel',
        documentation: ''
    },
    {
        label: 'vbCr',
        documentation: ''
    },
    {
        label: 'vbCritical',
        documentation: ''
    },
    {
        label: 'vbCrLf',
        documentation: ''
    },
    {
        label: 'vbCurrency',
        documentation: ''
    },
    
    {
        label: 'vbDatabaseCompare',
        documentation: ''
    },
    {
        label: 'vbDataObject',
        documentation: ''
    },
    {
        label: 'vbDate',
        documentation: ''
    },
    {
        label: 'vbDecimal',
        documentation: ''
    },
    {
        label: 'vbDouble',
        documentation: ''
    },
    {
        label: 'vbEmpty',
        documentation: ''
    },
    {
        label: 'vbError',
        documentation: ''
    },
    {
        label: 'vbExclamation',
        documentation: ''
    },
    {
        label: 'vbFalse',
        documentation: ''
    },
    {
        label: 'vbFirstFourDays',
        documentation: ''
    },
    {
        label: 'vbFirstFullWeek',
        documentation: ''
    },
    {
        label: 'vbFormFeed',
        documentation: ''
    },
    {
        label: 'vbFriday',
        documentation: ''
    },
    {
        label: 'vbGeneralDate',
        documentation: ''
    },
    
    {
        label: 'vbIgnore',
        documentation: ''
    },
    {
        label: 'vbInformation',
        documentation: ''
    },
    {
        label: 'vbInteger',
        documentation: ''
    },
    {
        label: 'vbLf',
        documentation: ''
    },
    {
        label: 'vbLong',
        documentation: ''
    },
    {
        label: 'vbLongDate',
        documentation: ''
    },
    {
        label: 'vbLongTime',
        documentation: ''
    },
    {
        label: 'vbMonday',
        documentation: ''
    },
    {
        label: 'vbMsgBoxHelpButton',
        documentation: ''
    },
    {
        label: 'vbMsgBoxRight',
        documentation: ''
    },
    {
        label: 'vbMsgBoxRtlReading',
        documentation: ''
    },
    {
        label: 'vbNewLine',
        documentation: ''
    },
    {
        label: 'vbNo',
        documentation: ''
    },
    {
        label: 'vbNull',
        documentation: ''
    },
    {
        label: 'vbNullChar',
        documentation: ''
    },
    {
        label: 'vbNullString',
        documentation: ''
    },
    {
        label: 'vbObject',
        documentation: ''
    },
    {
        label: 'vbObjectError',
        documentation: ''
    },
    {
        label: 'vbOK',
        documentation: ''
    },
    {
        label: 'vbOKCancel',
        documentation: ''
    },
    {
        label: 'vbOKOnly',
        documentation: ''
    },
    {
        label: 'vbQuestion',
        documentation: ''
    },
    
    {
        label: 'vbRetry',
        documentation: ''
    },
    {
        label: 'vbRetryCancel',
        documentation: ''
    },
    {
        label: 'vbSaturday',
        documentation: ''
    },
    {
        label: 'vbShortDate',
        documentation: ''
    },
    {
        label: 'vbShortTime',
        documentation: ''
    },
    {
        label: 'vbSingle',
        documentation: ''
    },
    {
        label: 'vbString',
        documentation: ''
    },
    {
        label: 'vbSunday',
        documentation: ''
    },
    {
        label: 'vbSystemModal',
        documentation: ''
    },
    {
        label: 'vbTab',
        documentation: ''
    },
    {
        label: 'vbTextCompare',
        documentation: ''
    },
    {
        label: 'vbThursday',
        documentation: ''
    },
    {
        label: 'vbTrue',
        documentation: ''
    },
    {
        label: 'vbTuesday',
        documentation: ''
    },
    {
        label: 'vbUseDefault',
        documentation: ''
    },
    {
        label: 'vbUseSystem',
        documentation: ''
    },
    {
        label: 'vbUseSystemDayOfWeek',
        documentation: ''
    },
    {
        label: 'vbVariant',
        documentation: ''
    },
    {
        label: 'vbVerticalTab',
        documentation: ''
    },
    {
        label: 'vbWednesday',
        documentation: ''
    },
    {
        label: 'vbYes',
        documentation: ''
    },
    {
        label: 'vbYesNo',
        documentation: ''
    },
    {
        label: 'vbYesNoCancel',
        documentation: ''
    },
]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Constant
    i.detail = 'Constant'
    // i.insertText = new SnippetString(i.insertText)
}

for (var i of colorItems) {
    i.kind = CompletionItemKind.Color
    i.detail = 'Constant (Color)'
    // i.insertText = new SnippetString(i.insertText)
}

export default [...items, ...colorItems];