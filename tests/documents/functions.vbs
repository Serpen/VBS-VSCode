

''' <summary>Returns the absolute value of a number.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Public Function Abs(expr)

End Function

''' <summary></summary>
Public Function Array()

End Function

''' <summary>Returns the unicode code of a character.</summary>
''' <param name="char">The character to get the code for. If a string is used, the code for the first character is given.</param>
Public Function Asc(char)

End Function


''' <summary>Returns the ANSI character code corresponding to the first letter in a string.</summary>
Public Function AscB(char)

End Function


''' <summary>Function that returns the Unicode (wide) character code that represents a specific Unicode character.</summary>
Public Function AscW(char)

End Function


''' <summary>Returns the arctangent of a number.</summary>
Public Function Atn(number)

End Function


''' <summary>Returns an expression that has been converted to a Variant of subtype Boolean.</summary>
Public Function CBool(expr)

End Function


''' <summary></summary>
Public Function CByte(expr)

End Function


''' <summary></summary>
Public Function CCur(expr)

End Function


''' <summary></summary>
Public Function CDate(expr)

End Function


''' <summary></summary>
Public Function CDbl(expr)

End Function


''' <summary></summary>
Public Function Chr(charcode)

End Function


''' <summary></summary>
Public Function ChrB(charcode)

End Function


''' <summary></summary>
Public Function ChrW(charcode)

End Function


''' <summary></summary>
Public Function CInt(expr)

End Function


''' <summary></summary>
Public Function CLng(expr)

End Function


''' <summary></summary>
Public Function Cos(number)

End Function


''' <summary></summary>
Public Function CreateObject(class)

End Function

''' <summary></summary>
Public Function CreateObject(class, location)

End Function


''' <summary></summary>
Public Function CSng(expr)

End Function


''' <summary></summary>
Public Function CStr(expr)

End Function


''' <summary></summary>
Public Property Get Date

End Property


''' <summary>Returns a date to which a specified time interval has been added.</summary>
''' <param name="interval">String expression that is the interval you want to add. See Settings section for values.</param>
''' <param name="number">Numeric expression that is the number of interval you want to add. The numeric expression can either be positive, for dates in the future, or negative, for dates in the past.</param>
''' <param name="date">Variant or literal representing the date to which interval is added.</param>
Public Function DateAdd(interval, number, date)

End Function


''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">Required. String expression that is the interval you want to use to calculate the differences between date1 and date2. See Settings section for values.</param>
''' <param name="date1">Required. Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Required. Date expressions. Two dates you want to use in the calculation..</param>
Public Function DateDiff(interval, date1, date2)

End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">Required. String expression that is the interval you want to use to calculate the differences between date1 and date2. See Settings section for values.</param>
''' <param name="date1">Required. Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Required. Date expressions. Two dates you want to use in the calculation..</param>
''' <param name="firstdayofweek">Optional. Constant that specifies the day of the week. If not specified, Sunday is assumed. See Settings section for values.</param>
Public Function DateDiff(interval, date1, date2, firstdayofweek)

End Function

''' <summary>Returns the number of intervals between two dates.</summary>
''' <param name="interval">Required. String expression that is the interval you want to use to calculate the differences between date1 and date2. See Settings section for values.</param>
''' <param name="date1">Required. Date expressions. Two dates you want to use in the calculation.</param>
''' <param name="date2">Required. Date expressions. Two dates you want to use in the calculation..</param>
''' <param name="firstdayofweek">Optional. Constant that specifies the day of the week. If not specified, Sunday is assumed. See Settings section for values.</param>
''' <param name="firstweekofyear">Optional. Constant that specifies the first week of the year. If not specified, the first week is assumed to be the week in which January 1 occurs. See Settings section for values.</param>
Public Function DateDiff(interval, date1, date2, firstdayofweek, firstweekofyear)

End Function


''' <summary></summary>
Public Function DatePart(interval, date)

End Function

''' <summary></summary>
Public Function DatePart(interval, date, firstdayofweek)

End Function

''' <summary></summary>
Public Function DatePart(interval, date, firstdayofweek, firstweekofyear)

End Function


''' <summary></summary>
Public Function DateSerial(year, month, day)

End Function


''' <summary></summary>
Public Function DateValue(date)

End Function


''' <summary></summary>
Public Function Day(date)

End Function

Public Class Debug

    ''' <summary></summary>
    Public Sub Write()

    End Sub


    ''' <summary></summary>
    Public Sub WriteLine()

    End Sub
End Class


Public Function Double()

End Function

Public Class Err

    ''' <summary></summary>
    Public Sub Clear()

    End Sub


    ''' <summary></summary>
    Public Property Get Description

    End Property


    ''' <summary></summary>
    Public Property Get HelpContext

    End Property


    ''' <summary></summary>
    Public Property Get HelpFile

    End Property


    ''' <summary></summary>
    Public Property Get Number

    End Property


     ''' <summary></summary>
    Public Sub Raise(number, source, description, helpfile, helpcontext)

    End Sub


    ''' <summary></summary>
    Public Property Get Source

    End Property
End Class


''' <summary></summary>
Public Function Escape(string)

End Function


''' <summary></summary>
Public Function Eval(expr)

End Function


''' <summary></summary>
Public Function Exp(number)

End Function


''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
Public Function Filter(InputStrings, Value)

End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
Public Function Filter(InputStrings, Value, Include)

End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
Public Function Filter(InputStrings, Value, Include, Compare)

End Function


''' <summary></summary>
Public Function Fix(number)

End Function


''' <summary></summary>
Public Function FormatCurrency(Expression)

End Function

''' <summary></summary>
Public Function FormatCurrency(Expression,NumDigitsAfterDecimal)

End Function

''' <summary></summary>
Public Function FormatCurrency(Expression,NumDigitsAfterDecimal ,IncludeLeadingDigit)

End Function

''' <summary></summary>
Public Function FormatCurrency(Expression,NumDigitsAfterDecimal ,IncludeLeadingDigit ,UseParensForNegativeNumbers)

End Function

''' <summary></summary>
Public Function FormatCurrency(Expression,NumDigitsAfterDecimal ,IncludeLeadingDigit ,UseParensForNegativeNumbers ,GroupDigits)

End Function


''' <summary></summary>
Public Function FormatDateTime(Date)

End Function

''' <summary></summary>
Public Function FormatDateTime(Date, NamedFormat)

End Function


''' <summary></summary>
Public Function FormatNumber(Expression)

End Function

''' <summary></summary>
Public Function FormatNumber(Expression ,NumDigitsAfterDecimal)

End Function

''' <summary></summary>
Public Function FormatNumber(Expression ,NumDigitsAfterDecimal ,IncludeLeadingDigit)

End Function

''' <summary></summary>
Public Function FormatNumber(Expression ,NumDigitsAfterDecimal ,IncludeLeadingDigit ,UseParensForNegativeNumbers)

End Function

''' <summary></summary>
Public Function FormatNumber(Expression ,NumDigitsAfterDecimal ,IncludeLeadingDigit ,UseParensForNegativeNumbers ,GroupDigits)

End Function


''' <summary></summary>
Public Function FormatPercent(Expression)

End Function

''' <summary></summary>
Public Function FormatPercent(Expression ,NumDigitsAfterDecimal)

End Function

''' <summary></summary>
Public Function FormatPercent(Expression ,NumDigitsAfterDecimal ,IncludeLeadingDigit)

End Function

''' <summary></summary>
Public Function FormatPercent(Expression ,NumDigitsAfterDecimal ,IncludeLeadingDigit ,UseParensForNegativeNumbers)

End Function

''' <summary></summary>
Public Function FormatPercent(Expression ,NumDigitsAfterDecimal ,IncludeLeadingDigit ,UseParensForNegativeNumbers ,GroupDigits)

End Function


''' <summary></summary>
Public Property Get GetLocale

End Property


''' <summary>?</summary>
Public Function GetObject(pathname)

End Function

''' <summary>?</summary>
Public Function GetObject(pathname, class)

End Function


''' <summary></summary>
Public Function GetRef(procname)

End Function


''' <summary></summary>
Public Property Get GetUILanguage

End Property


''' <summary></summary>
Public Function Hex(number)

End Function


''' <summary>Returns a whole number between 0 and 23, inclusive, representing the hour of the day.</summary>
Public Function Hour(time)

End Function


''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Public Function InputBox(prompt As String)

End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Public Function InputBox(prompt As String, title As String)

End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Public Function InputBox(prompt As String, title As String, default As String)

End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Public Function InputBox(prompt As String, title As String, default As String, xpos)

End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Public Function InputBox(prompt As String, title As String, default As String, xpos, ypos)

End Function

''' <summary>Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns the contents of the text box.</summary>
Public Function InputBox(prompt As String, title As String, default As String, xpos, ypos, helpfile, context)

End Function


''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Public Function InStr(string1, string2)

End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Public Function InStr(start, string1, string2)

End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Public Function InStr(string1, string2 compare)

End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Public Function InStr(start,string1, string2, compare)

End Function


''' <summary></summary>
Public Function InStrB()

End Function


''' <summary></summary>
Public Function InStrRev()

End Function


''' <summary></summary>
Public Function Int()

End Function


''' <summary></summary>
Public Function IsArray()

End Function


''' <summary></summary>
Public Function IsDate()

End Function


''' <summary></summary>
Public Function IsEmpty()

End Function


''' <summary></summary>
Public Function IsNull()

End Function


''' <summary>Returns a Boolean value indicating whether an expression can be evaluated as a number.</summary>
Public Function IsNumeric(expr)

End Function


''' <summary></summary>
Public Function IsObject()

End Function


''' <summary></summary>
Public Function Join()

End Function


''' <summary></summary>
Public Function LBound()

End Function


''' <summary></summary>
Public Function LCase()

End Function


''' <summary></summary>
Public Function Left()

End Function


''' <summary></summary>
Public Function LeftB()

End Function


''' <summary></summary>
Public Function Len()

End Function


''' <summary></summary>
Public Function LenB()

End Function


''' <summary></summary>
Public Function LoadPicture()

End Function


''' <summary></summary>
Public Function Log()

End Function


''' <summary></summary>
Public Function LTrim()

End Function

Public Class Match
    ''' <summary></summary>
    Public Property Get FirstIndex

    End Property


    ''' <summary></summary>
    Public Property Get Length

    End Property


    ''' <summary></summary>
    Public Property Get Value

    End Property
End Class

''' <summary></summary>
Public Function Me()

End Function


''' <summary>Returns a specified number of characters from a string.</summary>
Public Function Mid(string, start)

End Function

''' <summary>Returns a specified number of characters from a string.</summary>
Public Function Mid(string, start, length)

End Function


''' <summary></summary>
Public Function MidB()

End Function


''' <summary></summary>
Public Function Minute()

End Function


''' <summary></summary>
Public Function Month()

End Function


''' <summary></summary>
Public Function MonthName()

End Function


''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
Public Function MsgBox(prompt)

End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. See Settings section for values. If omitted, the default value for buttons is 0.</param>
Public Function MsgBox(prompt, buttons)

End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. See Settings section for values. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
Public Function MsgBox(prompt, buttons, title)

End Function

''' <summary>Displays a message in a dialog box, waits for the user to click a button, and returns a value indicating which button the user clicked.</summary>
''' <param name="prompt">String expression displayed as the message in the dialog box</param>
''' <param name="buttons">Numeric expression that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. See Settings section for values. If omitted, the default value for buttons is 0.</param>
''' <param name="title">String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.</param>
''' <param name="helpfile">String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If helpfile is provided, context must also be provided. Not available on 16-bit platforms.</param>
''' <param name="context">Numeric expression that identifies the Help context number assigned by the Help author to the appropriate Help topic. If context is provided, helpfile must also be provided. Not available on 16-bit platforms.</param>
Public Function MsgBox(prompt, buttons, title, helpfile, context)

End Function


''' <summary>Returns the current date and time according to the setting of your computer's system date and time.)</summary>
Public Property Get Now

End Property


''' <summary></summary>
Public Function Oct()

End Function

Class RegExp
    ''' <summary></summary>
    Public Function Execute()

    End Function


    ''' <summary></summary>
    Public Property Get Global

    End Property


     ''' <summary></summary>
    Public Property Get IgnoreCase

    End Property


    ''' <summary></summary>
    Public Function Replace()

    End Function


    ''' <summary></summary>
    Public Function Test()

    End Function

End Class


''' <summary></summary>
Public Function Replace()

End Function


''' <summary>Returns a whole number representing an RGB color value.</summary>
Public Function RGB(red, green, blue)

End Function


''' <summary></summary>
Public Function Right()

End Function


''' <summary></summary>
Public Function RightB()

End Function


''' <summary></summary>
Public Function Rnd()

End Function


''' <summary></summary>
Public Function Round()

End Function


''' <summary></summary>
Public Function RTrim()

End Function


''' <summary></summary>
Public Property Get ScriptEngine

End Property


''' <summary></summary>
Public Property Get ScriptEngineBuildVersion

End Property


''' <summary></summary>
Public Property Get ScriptEngineMajorVersion

End Property


''' <summary></summary>
Public Property Get ScriptEngineMinorVersion

End Property


''' <summary></summary>
Public Function Second()

End Function


''' <summary>undocumented</summary>
Public Function SetLocale(integer)

End Function


''' <summary></summary>
Public Function Sgn()

End Function


''' <summary></summary>
Public Function Sin()

End Function


''' <summary></summary>
Public Function Space()

End Function


''' <summary></summary>
Public Function Split(expression)

End Function

''' <summary></summary>
Public Function Split(expression, delimiter)

End Function

''' <summary></summary>
Public Function Split(expression, delimiter, count)

End Function

''' <summary></summary>
Public Function Split(expression, delimiter, count, compare)

End Function


''' <summary></summary>
Public Function Sqr()

End Function


''' <summary></summary>
Public Function StrComp()

End Function


''' <summary></summary>
Public Function StrReverse()

End Function


''' <summary></summary>
Public Function Tan()

End Function



''' <summary></summary>
Public Property Get Time

End Property


''' <summary></summary>
Public Property Get Timer

End Property


''' <summary></summary>
Public Function TimeSerial()

End Function


''' <summary></summary>
Public Function TimeValue()

End Function


''' <summary></summary>
Public Function Trim()

End Function


''' <summary></summary>
Public Function TypeName()

End Function


''' <summary>Returns the largest available subscript for the indicated dimension of an array.</summary>
Public Function UBound()

End Function


''' <summary>Returns a string that has been converted to uppercase.</summary>
Public Function UCase(str)

End Function


''' <summary></summary>
Public Function Unescape()

End Function


''' <summary></summary>
Public Function VarType()

End Function


''' <summary></summary>
Public Function Weekday()

End Function


''' <summary></summary>
Public Function WeekdayName()

End Function

Public Class WScript

    ''' <summary></summary>
    Public Function ConnectObject()

    End Function


    ''' <summary></summary>
    Public Function CreateObject()

    End Function


    ''' <summary></summary>
    Public Function DisconnectObject()

    End Function


    ''' <summary></summary>
    Public Function Echo()

    End Function


    ''' <summary></summary>
    Public Function GetObject()

    End Function


    ''' <summary></summary>
    Public Function Quit()

    End Function


    ''' <summary></summary>
    Public Function Sleep()

    End Function


    ''' <summary></summary>
    Public Property Get Arguments

    End Property


    ''' <summary></summary>
    Public Property Get BuildVersion

    End Property


    ''' <summary></summary>
    Public Property Get FullName

    End Property


    ''' <summary></summary>
    Public Property Get Interactive

    End Property


    ''' <summary></summary>
    Public Property Get Name

    End Property


    ''' <summary></summary>
    Public Property Get Path

    End Property


    ''' <summary></summary>
    Public Property Get ScriptFullName

    End Property


    ''' <summary></summary>
    Public Property Get ScriptName

    End Property


    ''' <summary></summary>
    Public Property Get TimeOut

    End Property


    ''' <summary></summary>
    Public Property Get Version

    End Property

End Class

''' <summary>Returns a whole number representing the year.</summary>
''' <param name="date">Any expression that can represent a date</param>
Public Function Year(date)

End Function

Public Const False = Empty ' Boolean
Public Const True = Empty
Public Const vbAbort = Empty
Public Const vbAbortRetryIgnore = Empty
Public Const vbApplicationModal = Empty
Public Const vbArray = Empty
Public Const vbBinaryCompare = Empty ' Perform a binary comparison
Public Const vbBoolean = Empty
Public Const vbByte = Empty
Public Const vbCancel = Empty
Public Const vbCr = Empty
Public Const vbCritical = Empty
Public Const vbCrLf = Empty
Public Const vbCurrency = Empty
Public Const vbDatabaseCompare = Empty
Public Const vbDataObject = Empty
Public Const vbDate = Empty
Public Const vbDecimal = Empty
Public Const vbDefaultButton1 = Empty
Public Const vbDefaultButton2 = Empty
Public Const vbDefaultButton3 = Empty
Public Const vbDefaultButton4 = Empty
Public Const vbEmpty = Empty
Public Const vbError = Empty
Public Const vbExclamation = Empty
Public Const vbFalse = Empty
Public Const vbFirstFourDays = Empty
Public Const vbFirstFullWeek = Empty
Public Const vbFirstJan1 = Empty
Public Const vbFormFeed = Empty
Public Const vbFriday = Empty
Public Const vbGeneralDate = Empty
Public Const vbIgnore = Empty
Public Const vbInformation = Empty
Public Const vbInteger = Empty
Public Const vbLf = Empty
Public Const vbLong = Empty
Public Const vbLongDate = Empty
Public Const vbLongTime = Empty
Public Const vbMonday = Empty
Public Const vbMsgBoxHelpButton = Empty
Public Const vbMsgBoxRight = Empty
Public Const vbMsgBoxRtlReading = Empty
Public Const vbNewLine = Empty
Public Const vbNo = Empty
Public Const vbNull = Empty
Public Const vbNullChar = Empty
Public Const vbNullString = Empty
Public Const vbObject = Empty
Public Const vbObjectError = Empty
Public Const vbOK = Empty
Public Const vbOKCancel = Empty ' Display OK and Cancel buttons
Public Const vbOKOnly = Empty ' Display OK button only.
Public Const vbQuestion = Empty ' Display Warning Query icon.
Public Const vbRetry = Empty ' Retry button was clicked
Public Const vbRetryCancel = Empty ' Display Retry and Cancel buttons.
Public Const vbSaturday = Empty
Public Const vbShortDate = Empty
Public Const vbShortTime = Empty
Public Const vbSingle = Empty
Public Const vbString = Empty
Public Const vbSunday = Empty
Public Const vbSystemModal = Empty
Public Const vbTab = Empty
Public Const vbTextCompare = Empty ' Perform a textual comparison
Public Const vbThursday = Empty
Public Const vbTrue = Empty
Public Const vbTuesday = Empty
Public Const vbUseDefault = Empty
Public Const vbUseSystem = Empty
Public Const vbUseSystemDayOfWeek = Empty
Public Const vbVariant = Empty
Public Const vbVerticalTab = Empty
Public Const vbWednesday = Empty
Public Const vbYes = Empty
Public Const vbYesNo = Empty
Public Const vbYesNoCancel = Empty

Public Const vbBlack = &h00
Public Const vbBlue = &hFF0000
Public Const vbCyan = &hFFFF00
Public Const vbGreen = &hFF00
Public Const vbMagenta = &hFF00FF
Public Const vbRed = &hFF
Public Const vbWhite = &hFFFFFF
Public Const vbYellow = &hFFFF
