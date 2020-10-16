''' <summary>Returns the absolute value of a number.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function Abs(expr)
End Function

''' <summary>Returns a Variant containing an array.</summary>
''' <param name="arglist">arglist argument is a comma-delimited list of values that are assigned to the elements of an array</param>
Function Array(arglist)
End Function

''' <summary>Returns the unicode code of a character.</summary>
''' <param name="char">The character to get the code for. If a string is used, the code for the first character is given.</param>
Function Asc(char)
End Function

''' <summary>Returns the ANSI character code corresponding to the first letter in a string.</summary>
Function AscB(char)
End Function

''' <summary>Function that returns the Unicode (wide) character code that represents a specific Unicode character.</summary>
Function AscW(char)
End Function

''' <summary>Returns the arctangent of a number.</summary>
Function Atn(number)
End Function

''' <summary>Returns an expression that has been converted to a Variant of subtype Boolean.</summary>
Function CBool(expr)
End Function

Function CByte(expr)
End Function

Function CCur(expr)
End Function

Function CDate(expr)
End Function

Function CDbl(expr)
End Function

Function Chr(charcode)
End Function

Function ChrB(charcode)
End Function

Function ChrW(charcode)
End Function

Function CInt(expr)
End Function

Function CLng(expr)
End Function

Function Cos(number)
End Function

Function CreateObject(classname)
End Function

Function CreateObject(classname, location)
End Function

Function CSng(expr)
End Function

Function CStr(expr)
End Function

Property Get Date
End Property

''' <summary>Returns a date to which a specified time interval has been added.</summary>
''' <param name="interval">String expression that is the interval you want to add. See Settings section for values.</param>
''' <param name="number">Numeric expression that is the number of interval you want to add. The numeric expression can either be positive, for dates in the future, or negative, for dates in the past.</param>
''' <param name="date">Variant or literal representing the date to which interval is added.</param>
Function DateAdd(interval, number, date)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2. See Settings section for values.</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
Function DateDiff(interval, date1, date2)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2. See Settings section for values.</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed. See Settings section for values.</param>
Function DateDiff(interval, date1, date2, firstdayofweek)
End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">String expression that is the interval you want to use to calculate the differences between date1 and date2. See Settings section for values.</param>
''' <param name="date1">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed. See Settings section for values.</param>
''' <param name="firstweekofyear">Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs. See Settings section for values.</param>
Function DateDiff(interval, date1, date2, firstdayofweek, firstweekofyear)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
Function DatePart(interval, date)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
Function DatePart(interval, date, firstdayofweek)
End Function

''' <summary>Returns the specified part of a given date.</summary>
''' <param name="interval">String expression that is the interval of time you want to return</param>
''' <param name="date">Date expression you want to evaluate</param>
''' <param name="firstdayofweek">Constant that specifies the day of the week. If not specified, Sunday is assumed</param>
''' <param name="firstweekofyear">Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs</param>
Function DatePart(interval, date, firstdayofweek, firstweekofyear)
End Function

''' <summary>Returns a Variant of subtype Date for a specified year, month, and day.</summary>
''' <param name="year">Number between 100 and 9999, inclusive, or a numeric expression.</param>
Function DateSerial(year, month, day)
End Function

''' <summary>Returns a Variant of subtype Date.</summary>
Function DateValue(date)
End Function

''' <summary>Returns a whole number between 1 and 31, inclusive, representing the day of the month.</summary>
Function Day(date)
End Function

''' <summary>Returns</summary>
Function Escape(str)
End Function

''' <summary>Evaluates an expression and returns the result.</summary>
Function Eval(expr)
End Function

''' <summary>Returns e (the base of natural logarithms) raised to a power.</summary>
Function Exp(number)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
Function Filter(InputStrings, Value)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
''' <param name="Include">Boolean value indicating whether to return substrings that include or exclude Value. If Include is True, Filter returns the subset of the array that contains Value as a substring. If Include is False, Filter returns the subset of the array that does not contain Value as a substring.</param>
Function Filter(InputStrings, Value, Include)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
''' <param name="InputStrings">One-dimensional array of strings to be searched.</param>
''' <param name="Value">String to search for.</param>
''' <param name="Include">Boolean value indicating whether to return substrings that include or exclude Value. If Include is True, Filter returns the subset of the array that contains Value as a substring. If Include is False, Filter returns the subset of the array that does not contain Value as a substring.</param>
''' <param name="Compare">Numeric value indicating the kind of string comparison to use</param>
Function Filter(InputStrings, Value, Include, Compare)
End Function

Function Fix(number)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
Function FormatCurrency(Expression)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers)
End Function

''' <summary>Returns an expression formatted as a currency value using the currency symbol defined in the system control panel.</summary>
''' <param name="Expression">Expression to be formatted.</param>
''' <param name="NumDigitsAfterDecimal">Numeric value indicating how many places to the right of the decimal are displayed. Default value is -1, which indicates that the computer's regional settings are used</param>
''' <param name="IncludeLeadingDigit">Tristate constant that indicates whether or not a leading zero is displayed for fractional values</param>
''' <param name="UseParensForNegativeNumbers">Tristate constant that indicates whether or not to place negative values within parentheses</param>
''' <param name="GroupDigits">Tristate constant that indicates whether or not numbers are grouped using the group delimiter specified in the computer's regional settings</param>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
End Function

Function FormatDateTime(Date)
End Function

Function FormatDateTime(Date, NamedFormat)
End Function

Function FormatNumber(Expression)
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal)
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit)
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers)
End Function

Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
End Function

Function FormatPercent(Expression)
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal)
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit)
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers)
End Function

Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
End Function

Property Get GetLocale
End Property

''' <summary>?</summary>
Function GetObject(pathname)
End Function

''' <summary>?</summary>
Function GetObject(pathname, classname)
End Function

Function GetRef(procname)
End Function

Property Get GetUILanguage
End Property

Function Hex(number)
End Function

''' <summary>Returns a whole number between 0 and 23, inclusive, representing the hour of the day.</summary>
Function Hour(time)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default, xpos)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default, xpos, ypos)
End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Function InputBox(prompt, title, default, xpos, ypos, helpfile, context)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
Function InStr(string1, string2)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search</param>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
Function InStr(start, string1, string2)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
''' <param name="compare">Specifies the string comparison to use</param>
Function InStr(string1, string2, compare)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
''' <param name="start">Specifies the starting position for search</param>
''' <param name="string1">The string to be searched</param>
''' <param name="string2">The string expression to search for</param>
''' <param name="compare">Specifies the string comparison to use</param>
Function InStr(start, string1, string2, compare)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(string1, string2)
End Function
''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(start, string1, string2)
End Function
''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(string1, string2, compare)
End Function
''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStrB(start, string1, string2, compare)
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
Function InStrRev(string1, string2)
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
Function InStrRev(string1, string2, start)
End Function

''' <summary>Returns the position of an occurrence of one string within another, from the end of string.</summary>
Function InStrRev(string1, string2, start, compare)
End Function

Function Int(number)
End Function

Function IsArray(var)
End Function

Function IsDate(expr)
End Function

Function IsEmpty(expr)
End Function

Function IsNull(expr)
End Function

''' <summary>Returns a Boolean value indicating whether an expression can be evaluated as a number.</summary>
Function IsNumeric(expr)
End Function

Function IsObject(expr)
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
Function Join(list)
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
Function Join(list, delimiter)
End Function

Function LBound(arrayname)
End Function

Function LBound(arrayname, dimension)
End Function

Function LCase(str)
End Function

Function Left(str, length)
End Function

Function LeftB(str, length)
End Function

''' <summary>Returns the number of characters in a string or the number of bytes required to store a variable.</summary>
Function Len(str)
End Function

Function LenB(str)
End Function

Function LoadPicture(picturename)
End Function

Function Log(number)
End Function

Function LTrim(str)
End Function

Function Me()
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function Mid(str, start)
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function Mid(str, start, length)
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function MidB(str, start)
End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Function MidB(str, start, length)
End Function

Function Minute(time)
End Function

Function Month(date)
End Function

Function MonthName(date)
End Function

Function MonthName(date, abbrevation)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
Function MsgBox(prompt)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. See Settings section for values. If omitted, the default value for buttons is 0.</param>
Function MsgBox(prompt, buttons)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. See Settings section for values. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
Function MsgBox(prompt, buttons, title)
End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. See Settings section for values. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
''' <param name="helpfile">String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If helpfile is provided, context must also be provided. Not available on 16-bit platforms.</param>
''' <param name="context">Numeric expression that identifies the Help context number assigned by the Help author to the appropriate Help topic. If context is provided, helpfile must also be provided. Not available on 16-bit platforms.</param>
Function MsgBox(prompt, buttons, title, helpfile, context)
End Function

''' <summary>Returns the current date and time according to the setting of your computer's system date and time.</summary>
Property Get Now
End Property

Function Oct(number)
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
Function Replace(str, find, replacewith)
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
Function Replace(str, find, replacewith, start)
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
Function Replace(str, find, replacewith, start, count)
End Function

''' <summary>Returns a string in which a specified substring has been replaced with another substring a specified number of times.</summary>
Function Replace(str, find, replacewith, start, count, compare)
End Function

''' <summary>Returns a whole number representing an RGB color value.</summary>
Function RGB(red, green, blue)
End Function

Function Right(str, length)
End Function

Function RightB(str, length)
End Function

Function Rnd()
End Function

Function Rnd(number)
End Function

Function Round(number, digits)
End Function

Function RTrim(str)
End Function

Property Get ScriptEngine
End Property

Property Get ScriptEngineBuildVersion
End Property

Property Get ScriptEngineMajorVersion
End Property

Property Get ScriptEngineMinorVersion
End Property

Function Second(time)
End Function

''' <summary>undocumented</summary>
Function SetLocale(integer)
End Function

Function Sgn(number)
End Function

Function Sin(number)
End Function

Function Space(number)
End Function

Function Split(str)
End Function

Function Split(str, delimiter)
End Function

Function Split(str, delimiter, count)
End Function

Function Split(str, delimiter, count, compare)
End Function

Function Sqr(number)
End Function

Function StrComp(string1, string2)
End Function

Function StrComp(string1, string2, compare)
End Function

Function StrReverse(str)
End Function

Function Tan(number)
End Function

Property Get Time
End Property

Property Get Timer
End Property

Function TimeSerial(hour, minute, second)
End Function

Function TimeValue(time)
End Function

Function Trim(str)
End Function

Function TypeName(var)
End Function

''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
Function UBound(arrayname)
End Function

''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
Function UBound(arrayname, dimension)
End Function

''' <summary>Returns a string that has been converted to uppercase.</summary>
Function UCase(str)
End Function

Function Unescape(str)
End Function

Function VarType(var)
End Function

Function Weekday(date)
End Function

Function Weekday(date, firstdayofweek)
End Function

Function WeekdayName(weekday)
End Function

Function WeekdayName(weekday, abbreviate)
End Function

Function WeekdayName(weekday, abbreviate, firstdayofweek)
End Function

''' <summary>Returns a whole number representing the year.</summary>
''' <param name="date">Any expression that can represent a date</param>
Function Year(date)
End Function

Const vbEmpty = 0
Const vbNull = 1
Const vbInteger = 2
Const vbLong = 3
Const vbSingle = 4
Const vbDouble = 5
Const vbCurrency = 6
Const vbDate = 7
Const vbString = 8
Const vbObject = 9
Const vbError = 10
Const vbBoolean = 11
Const vbVariant = 12
Const vbDataObject = 13
Const vbDecimal = 14
Const vbByte = 17
Const vbArray = 8192

Const False = False ' Boolean
Const True = True
Const vbUseDefault = -2
Const vbTrue = -1
Const vbFalse = 0

Const vbBinaryCompare = 0 ' Perform a binary comparison
Const vbTextCompare = 1 ' Perform a textual comparison
Const vbDatabaseCompare = 2 ' Only in Access

Const vbCr = Chr(13)
Const vbCrLf = Chr(13) & Chr(10)
Const vbFormFeed = Chr(12)
Const vbLf = Chr(10)
Const vbNewLine = Chr(13) & Chr(10)
Const vbNullChar = Chr(0)
Const vbNullString = Empty
Const vbTab = Chr(9)
Const vbVerticalTab = Chr(11)

Const vbGeneralDate = 0
Const vbLongDate = 1 ' Display a date using the long date format specified in your computer's regional settings.
Const vbShortDate = 2 ' Display a date using the short date format specified in your computer's regional settings.
Const vbLongTime = 3 ' Display a time using the time format specified in your computer's regional settings.
Const vbShortTime = 4

Const vbUseSystemDayOfWeek = 0
Const vbFirstJan1 = 1
Const vbFirstFourDays = 2
Const vbFirstFullWeek = 3

Const vbObjectError = &h80040000

Const vbMonday = 2
Const vbTuesday = 3
Const vbWednesday = 4
Const vbThursday = 5
Const vbFriday = 6
Const vbSaturday = 7
Const vbSunday = 1
Const vbUseSystem = 0

Const vbOKOnly = 0 ' Display OK button only.
Const vbOKCancel = 1 ' Display OK and Cancel buttons
Const vbAbortRetryIgnore = 2
Const vbYesNoCancel = 3
Const vbYesNo = 4
Const vbRetryCancel = 5 ' Display Retry and Cancel buttons.
Const vbCritical = 16
Const vbQuestion = 32 ' Display Warning Query icon.
Const vbExclamation = 48
Const vbInformation = 64
Const vbDefaultButton1 = 0
Const vbDefaultButton2 = 256
Const vbDefaultButton3 = 512
Const vbDefaultButton4 = 768
Const vbApplicationModal = 0
Const vbSystemModal = &h1000
Const vbMsgBoxHelpButton = &h4000
Const VbMsgBoxSetForeground = &h10000
Const vbMsgBoxRight = &h80000
Const vbMsgBoxRtlReading = &h100000

Const vbOK = 1
Const vbCancel = 2
Const vbAbort = 3
Const vbRetry = 4 ' Retry button was clicked
Const vbIgnore = 5
Const vbYes = 6
Const vbNo = 7

Const vbBlack = &h00
Const vbBlue = &hFF0000
Const vbCyan = &hFFFF00
Const vbGreen = &hFF00
Const vbMagenta = &hFF00FF
Const vbRed = &hFF
Const vbWhite = &hFFFFFF
Const vbYellow = &hFFFF

' Const SystemFolder = 1
' Const TemporaryFolder = 2
' Const WindowsFolder = 0