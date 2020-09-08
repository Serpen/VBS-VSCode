'use strict'
import { CompletionItem, CompletionItemKind } from "vscode";

var items : CompletionItem[] = [{
        label: 'Abs',
        documentation: 'Calculates the absolute value of a number.'
    },
    {
        label: 'ACos',
        documentation: 'Calculates the arcCosine of a number.'
    },
    {
        label: 'ASin',
        documentation: 'Calculates the arcsine of a number.'
    },
    {
        label: 'Assign',
        documentation: 'Assigns a variable by name with the data.'
    },
    {
        label: 'Asc',
        documentation: 'Returns the ASCII code of a character.'
    },
    {
        label: 'AscW',
        documentation: 'Returns the unicode code of a character.'
    },
    {
        label: 'ATan',
        documentation: 'Calculates the arctangent of a number.'
    },
    {
        label: 'Binary',
        documentation: 'Returns the binary representation of an expression.'
    },
    {
        label: 'BinaryLen',
        documentation: 'Returns the number of bytes in a binary variant.'
    },
    {
        label: 'BinaryMid',
        documentation: 'Extracts a number of bytes from a binary variant.'
    },
    {
        label: 'BinaryToString',
        documentation: 'Converts a binary variant into a string.'
    },
    {
        label: 'Call',
        documentation: 'Calls a user-defined or built-in function contained in first parameter.'
    },
    {
        label: 'Ceiling',
        documentation: 'Returns a number rounded up to the next integer.'
    },
    {
        label: 'Chr',
        documentation: 'Returns a character corresponding to an ASCII code.'
    },
    {
        label: 'ChrW',
        documentation: 'Returns a character corresponding to a unicode code.'
    },
    {
        label: 'Cos',
        documentation: 'Calculates the cosine of a number.'
    },
    {
        label: 'Dec',
        documentation: 'Returns a numeric representation of a hexadecimal string.'
    },
    {
        label: 'Eval',
        documentation: 'Return the value of the variable defined by a string. '
    },
    {
        label: 'Execute',
        documentation: 'Execute an expression. '
    },
    {
        label: 'Exp',
        documentation: 'Calculates e to the power of a number. '
    },
    {
        label: 'Hex',
        documentation: 'Returns a string representation of an integer or of a binary type converted to hexadecimal. '
    },
    {
        label: 'InputBox',
        documentation: 'Displays an input box to ask the user to enter a string. '
    },
    {
        label: 'Int',
        documentation: 'Returns the integer (whole number) representation of an expression. '
    },
    {
        label: 'IsAdmin',
        documentation: 'Checks if the current user has full administrator privileges. '
    },
    {
        label: 'IsArray',
        documentation: 'Checks if a variable is an array type. '
    },
    {
        label: 'IsBinary',
        documentation: 'Checks if a variable or expression is a binary type. '
    },
    {
        label: 'IsBool',
        documentation: 'Checks if a variable\'s base type is boolean. '
    },
    {
        label: 'IsDeclared',
        documentation: 'Check if a variable has been declared. '
    },
    {
        label: 'IsDllStruct',
        documentation: 'Checks if a variable is a DllStruct type. '
    },
    {
        label: 'IsFloat',
        documentation: 'Checks if the value of a variable or expression has a fractional component. '
    },
    {
        label: 'IsFunc',
        documentation: 'Checks if a variable or expression is a function type. '
    },
    {
        label: 'IsHWnd',
        documentation: 'Checks if a variable\'s base type is a pointer and window handle. '
    },
    {
        label: 'IsInt',
        documentation: 'Checks if the value of a variable or expression has no fractional component. '
    },
    {
        label: 'IsKeyword',
        documentation: 'Checks if a variable is a keyword (for example, Default). '
    },
    {
        label: 'IsNumber',
        documentation: 'Checks if a variable\'s base type is numeric. '
    },
    {
        label: 'IsObj',
        documentation: 'Checks if a variable or expression is an object type. '
    },
    {
        label: 'IsPtr',
        documentation: 'Checks if a variable\'s base type is a pointer. '
    },
    {
        label: 'IsString',
        documentation: 'Checks if a variable is a string type. '
    },
    {
        label: 'Log',
        documentation: 'Calculates the natural logarithm of a number. '
    },
    {
        label: 'MemGetStats',
        documentation: 'Retrieves memory related information. '
    },
    {
        label: 'Mod',
        documentation: 'Performs the modulus operation. '
    },
    {
        label: 'MsgBox',
        documentation: 'Displays a simple message box with optional timeout. '
    },
    {
        label: 'Number',
        documentation: 'Returns the numeric representation of an expression. '
    },
    {
        label: 'ObjCreate',
        documentation: 'Creates a reference to a COM object from the given classname. '
    },
    {
        label: 'ObjCreateInterface',
        documentation: 'Creates a reference to an object from the given classname/object pointer, interface identifier and description string. '
    },
    {
        label: 'ObjEvent',
        documentation: 'Handles incoming events from the given Object. '
    },
    {
        label: 'ObjGet',
        documentation: 'Retrieves a reference to a COM object from an existing process or filename. '
    },
    {
        label: 'ObjName',
        documentation: 'Returns the name or interface description of an Object. '
    },
    {
        label: 'Random',
        documentation: 'Generates a pseudo-random float-type number. '
    },
    {
        label: 'Round',
        documentation: 'Returns a number rounded to a specified number of decimal places. '
    },
    {
        label: 'Run',
        documentation: 'Runs an external program. '
    },
    {
        label: 'RunAs',
        documentation: 'Runs an external program under the context of a different user. '
    },
    {
        label: 'RunAsWait',
        documentation: 'Runs an external program under the context of a different user and pauses script execution until the program finishes. '
    },
    {
        label: 'RunWait',
        documentation: 'Runs an external program and pauses script execution until the program finishes. '
    },
    {
        label: 'Send',
        documentation: 'Sends simulated keystrokes to the active window. '
    },
    {
        label: 'SendKeepActive',
        documentation: 'Attempts to keep a specified window active during Send(). '
    },
    {
        label: 'SetError',
        documentation: 'Manually set the value of the @error macro (and optionally @extended, and "Return Value"). '
    },
    {
        label: 'SetExtended',
        documentation: 'Manually set the value of the @extended macro. '
    },
    {
        label: 'ShellExecute',
        documentation: 'Runs an external program using the ShellExecute API. '
    },
    {
        label: 'ShellExecuteWait',
        documentation: 'Runs an external program using the ShellExecute API and pauses script execution until it finishes. '
    },
    {
        label: 'Shutdown',
        documentation: 'Shuts down the system. '
    },
    {
        label: 'Sin',
        documentation: 'Calculates the sine of a number. '
    },
    {
        label: 'Sleep',
        documentation: 'Pause script execution. '
    },
    {
        label: 'SoundPlay',
        documentation: 'Play a sound file. '
    },
    {
        label: 'SoundSetWaveVolume',
        documentation: 'Sets the system wave volume by percent. '
    },
    {
        label: 'SplashImageOn',
        documentation: 'Creates a customizable image popup window. '
    },
    {
        label: 'SplashOff',
        documentation: 'Turns SplashText or SplashImage off. '
    },
    {
        label: 'SplashTextOn',
        documentation: 'Creates a customizable text popup window. '
    },
    {
        label: 'Sqrt',
        documentation: 'Calculates the square-root of a number. '
    },
    {
        label: 'SRandom',
        documentation: 'Set Seed for random number generation. '
    },
    {
        label: 'StatusbarGetText',
        documentation: 'Retrieves the text from a standard status bar control. '
    },
    {
        label: 'StderrRead',
        documentation: 'Reads from the STDERR stream of a previously run child process. '
    },
    {
        label: 'StdinWrite',
        documentation: 'Writes a number of characters to the STDIN stream of a previously run child process. '
    },
    {
        label: 'StdioClose',
        documentation: 'Closes all resources associated with a process previously run with STDIO redirection. '
    },
    {
        label: 'StdoutRead',
        documentation: 'Reads from the STDOUT stream of a previously run child process. '
    },
    {
        label: 'String',
        documentation: 'Returns the string representation of an expression. '
    },
    {
        label: 'StringAddCR',
        documentation: 'Takes a string and prefixes all linefeed characters ( Chr(10) ) with a carriage return character ( Chr(13) ). '
    },
    {
        label: 'StringCompare',
        documentation: 'Compares two strings with options. '
    },
    {
        label: 'StringFormat',
        documentation: 'Returns a formatted string (similar to the C sprintf() function). '
    },
    {
        label: 'StringFromASCIIArray',
        documentation: 'Converts an array of ASCII codes to a string. '
    },
    {
        label: 'StringInStr',
        documentation: 'Checks if a string contains a given substring. '
    },
    {
        label: 'StringIsAlNum',
        documentation: 'Checks if a string contains only alphanumeric characters. '
    },
    {
        label: 'StringIsAlpha',
        documentation: 'Checks if a string contains only alphabetic characters. '
    },
    {
        label: 'StringIsASCII',
        documentation: 'Checks if a string contains only ASCII characters in the range 0x00 - 0x7f (0 - 127). '
    },
    {
        label: 'StringIsDigit',
        documentation: 'Checks if a string contains only digit (0-9) characters. '
    },
    {
        label: 'StringIsFloat',
        documentation: 'Checks if a string is a floating point number. '
    },
    {
        label: 'StringIsInt',
        documentation: 'Checks if a string is an integer. '
    },
    {
        label: 'StringIsLower',
        documentation: 'Checks if a string contains only lowercase characters. '
    },
    {
        label: 'StringIsSpace',
        documentation: 'Checks if a string contains only whitespace characters. '
    },
    {
        label: 'StringIsUpper',
        documentation: 'Checks if a string contains only uppercase characters. '
    },
    {
        label: 'StringIsXDigit',
        documentation: 'Checks if a string contains only hexadecimal digit (0-9, A-F) characters. '
    },
    {
        label: 'StringLeft',
        documentation: 'Returns a number of characters from the left-hand side of a string. '
    },
    {
        label: 'StringLen',
        documentation: 'Returns the number of characters in a string. '
    },
    {
        label: 'StringLower',
        documentation: 'Converts a string to lowercase. '
    },
    {
        label: 'StringMid',
        documentation: 'Extracts a number of characters from a string. '
    },
    {
        label: 'StringRegExp',
        documentation: 'Check if a string fits a given regular expression pattern. '
    },
    {
        label: 'StringRegExpReplace',
        documentation: 'Replace text in a string based on regular expressions. '
    },
    {
        label: 'StringReplace',
        documentation: 'Replaces substrings in a string. '
    },
    {
        label: 'StringReverse',
        documentation: 'Reverses the contents of the specified string. '
    },
    {
        label: 'StringRight',
        documentation: 'Returns a number of characters from the right-hand side of a string. '
    },
    {
        label: 'StringSplit',
        documentation: 'Splits up a string into substrings depending on the given delimiters. '
    },
    {
        label: 'StringStripCR',
        documentation: 'Removes all carriage return values ( Chr(13) ) from a string. '
    },
    {
        label: 'StringStripWS',
        documentation: 'Strips the white space in a string. '
    },
    {
        label: 'StringToASCIIArray',
        documentation: 'Converts a string to an array containing the ASCII code of each character. '
    },
    {
        label: 'StringToBinary',
        documentation: 'Converts a string into binary data. '
    },
    {
        label: 'StringTrimLeft',
        documentation: 'Trims a number of characters from the left hand side of a string. '
    },
    {
        label: 'StringTrimRight',
        documentation: 'Trims a number of characters from the right hand side of a string. '
    },
    {
        label: 'StringUpper',
        documentation: 'Converts a string to uppercase. '
    },
    {
        label: 'Tan',
        documentation: 'Calculates the tangent of a number. '
    },
    {
        label: 'TimerDiff',
        documentation: 'Returns the difference in time from a previous call to TimerInit(). '
    },
    {
        label: 'TimerInit',
        documentation: 'Returns a handle that can be passed to TimerDiff() to calculate the difference in milliseconds. '
    },
    {
        label: 'ToolTip',
        documentation: 'Creates a tooltip anywhere on the screen. '
    },
    {
        label: 'UBound',
        documentation: 'Returns the size of array dimensions or the number of keys in a map. '
    },
    {
        label: 'VarGetType',
        documentation: 'Returns the internal type representation of a variant. '
    }
]

// Add the function icon and detail to each entry
for (var i of items) {
    i.kind = CompletionItemKind.Function
    i.detail = 'Function'
    // i.insertText = new SnippetString(i.insertText)
}

export default items;