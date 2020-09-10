'use strict'
import { CompletionItem, CompletionItemKind } from "vscode";

var items: CompletionItem[] = [
    {
        label: 'Abs',
        documentation: ''
    },
    {
        label: 'Asc',
        documentation: ''
    },
    {
        label: 'AscB',
        documentation: ''
    },
    {
        label: 'AscW',
        documentation: ''
    },
    {
        label: 'Atn',
        documentation: ''
    },
    {
        label: 'CBool',
        documentation: ''
    },
    {
        label: 'CByte',
        documentation: ''
    },
    {
        label: 'CCur',
        documentation: ''
    },
    {
        label: 'CDate',
        documentation: ''
    },
    {
        label: 'CDbl',
        documentation: ''
    },
    {
        label: 'Chr',
        documentation: ''
    },
    {
        label: 'ChrB',
        documentation: ''
    },
    {
        label: 'ChrW',
        documentation: ''
    },
    {
        label: 'CInt',
        documentation: ''
    },

    {
        label: 'CLng',
        documentation: ''
    },

    {
        label: 'Cos',
        documentation: ''
    },
    {
        label: 'CreateObject',
        documentation: ''
    },
    {
        label: 'CSng',
        documentation: ''
    },
    {
        label: 'CStr',
        documentation: ''
    },
    {
        label: 'DateAdd',
        documentation: ''
    },
    {
        label: 'DateDiff',
        documentation: ''
    },
    {
        label: 'DatePart',
        documentation: ''
    },
    {
        label: 'DateSerial',
        documentation: ''
    },
    {
        label: 'DateValue',
        documentation: ''
    },
    {
        label: 'Day',
        documentation: ''
    },
    {
        label: 'Double',
        documentation: ''
    },
    {
        label: 'Empty',
        documentation: ''
    },
    {
        label: 'Eqv',
        documentation: ''
    },
    {
        label: 'Err',
        documentation: 'Class'
    },
    {
        label: 'Erase',
        documentation: ''
    },
    {
        label: 'Escape',
        documentation: ''
    },
    {
        label: 'Eval',
        documentation: ''
    },
    {
        label: 'Execute',
        documentation: ''
    },
    {
        label: 'ExecuteGlobal',
        documentation: ''
    },
    {
        label: 'Exp',
        documentation: ''
    },
    {
        label: 'Filter',
        documentation: ''
    },
    {
        label: 'Fix',
        documentation: ''
    },
    {
        label: 'For',
        documentation: ''
    },
    {
        label: 'FormatCurrency',
        documentation: ''
    },
    {
        label: 'FormatDateTime',
        documentation: ''
    },
    {
        label: 'FormatNumber',
        documentation: ''
    },
    {
        label: 'FormatPercent',
        documentation: ''
    },
    {
        label: 'GetObject',
        documentation: ''
    },
    {
        label: 'GetRef',
        documentation: ''
    },
    {
        label: 'Hex',
        documentation: ''
    },
    {
        label: 'Hour',
        documentation: ''
    },
    {
        label: 'Imp',
        documentation: ''
    },
    {
        label: 'InputBox',
        documentation: ''
    },
    {
        label: 'InStr',
        documentation: ''
    },
    {
        label: 'InStrB',
        documentation: ''
    },
    {
        label: 'InStrRev',
        documentation: ''
    },
    {
        label: 'Int',
        documentation: ''
    },

    {
        label: 'IsArray',
        documentation: ''
    },
    {
        label: 'IsDate',
        documentation: ''
    },
    {
        label: 'IsEmpty',
        documentation: ''
    },
    {
        label: 'IsNull',
        documentation: ''
    },
    {
        label: 'IsNumeric',
        documentation: ''
    },
    {
        label: 'IsObject',
        documentation: ''
    },
    {
        label: 'Join',
        documentation: ''
    },
    {
        label: 'LBound',
        documentation: ''
    },
    {
        label: 'LCase',
        documentation: ''
    },
    {
        label: 'Left',
        documentation: ''
    },
    {
        label: 'LeftB',
        documentation: ''
    },
    {
        label: 'Len',
        documentation: ''
    },
    {
        label: 'LenB',
        documentation: ''
    },

    {
        label: 'LoadPicture',
        documentation: ''
    },
    {
        label: 'Log',
        documentation: ''
    },

    {
        label: 'LSet',
        documentation: ''
    },
    {
        label: 'LTrim',
        documentation: ''
    },
    {
        label: 'Mid',
        documentation: ''
    },
    {
        label: 'MidB',
        documentation: ''
    },
    {
        label: 'Minute',
        documentation: ''
    },
    {
        label: 'Month',
        documentation: ''
    },
    {
        label: 'MonthName',
        documentation: ''
    },
    {
        label: 'MsgBox',
        documentation: ''
    },
    {
        label: 'Oct',
        documentation: ''
    },

    {
        label: 'Randomize',
        documentation: ''
    },

    {
        label: 'Replace',
        documentation: ''
    },

    {
        label: 'RGB',
        documentation: ''
    },
    {
        label: 'Right',
        documentation: ''
    },
    {
        label: 'RightB',
        documentation: ''
    },
    {
        label: 'Rnd',
        documentation: ''
    },
    {
        label: 'Round',
        documentation: ''
    },
    {
        label: 'RSet',
        documentation: ''
    },
    {
        label: 'RTrim',
        documentation: ''
    },
    {
        label: 'Second',
        documentation: ''
    },
    {
        label: 'SetLocale',
        documentation: ''
    },
    {
        label: 'Sgn',
        documentation: ''
    },

    {
        label: 'Sin',
        documentation: ''
    },

    {
        label: 'Space',
        documentation: ''
    },
    {
        label: 'Split',
        documentation: ''
    },
    {
        label: 'Sqr',
        documentation: ''
    },

    {
        label: 'StrComp',
        documentation: ''
    },
    {
        label: 'StrReverse',
        documentation: ''
    },
    {
        label: 'Tan',
        documentation: ''
    },
    {
        label: 'TimeSerial',
        documentation: ''
    },
    {
        label: 'TimeValue',
        documentation: ''
    },
    {
        label: 'Trim',
        documentation: ''
    },
    {
        label: 'True',
        documentation: ''
    },
    {
        label: 'TypeName',
        documentation: ''
    },

    {
        label: 'UBound',
        documentation: ''
    },
    {
        label: 'UCase',
        documentation: ''
    },
    {
        label: 'Unescape',
        documentation: ''
    },
    {
        label: 'VarType',
        documentation: ''
    },
    {
        label: 'Weekday',
        documentation: ''
    },
    {
        label: 'WeekdayName',
        documentation: ''
    },
    {
        label: 'Year',
        documentation: ''
    },

]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Function
    i.detail = 'Function'
    // i.insertText = new SnippetString(i.insertText)
}

export default items;