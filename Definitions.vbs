''' <summary>Returns the absolute value of a number.</summary>
''' <param name="expr">Any valid numeric expression.</param>
Function Abs(expr)
End Function

''' <summary></summary>
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


''' <summary></summary>
Function CByte(expr)
End Function


''' <summary></summary>
Function CCur(expr)
End Function


''' <summary></summary>
Function CDate(expr)
End Function


''' <summary></summary>
Function CDbl(expr)
End Function


''' <summary></summary>
Function Chr(charcode)
End Function


''' <summary></summary>
Function ChrB(charcode)
End Function


''' <summary></summary>
Function ChrW(charcode)
End Function


''' <summary></summary>
Function CInt(expr)
End Function


''' <summary></summary>
Function CLng(expr)
End Function


''' <summary></summary>
Function Cos(number)
End Function


''' <summary></summary>
Function CreateObject(classname)
End Function

''' <summary></summary>
Function CreateObject(classname, location)
End Function


''' <summary></summary>
Function CSng(expr)
End Function


''' <summary></summary>
Function CStr(expr)
End Function


''' <summary></summary>
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


''' <summary></summary>
Function DatePart(interval, date)
End Function

''' <summary></summary>
Function DatePart(interval, date, firstdayofweek)
End Function

''' <summary></summary>
Function DatePart(interval, date, firstdayofweek, firstweekofyear)
End Function


''' <summary></summary>
Function DateSerial(year, month, day)
End Function


''' <summary></summary>
Function DateValue(date)
End Function


''' <summary></summary>
Function Day(date)
End Function


''' <summary>An intrinsic global object that can send output to a script debugger, such as the Microsoft Script Debugger.</summary>
Class Debug

	''' <summary></summary>
	Sub Write(str)
	End Sub
	
	''' <summary></summary>
	Sub WriteLine(str)
	End Sub
End Class

''' <summary>Object that stores data key, item pairs.</summary>
Class Dictionary
	
	Sub Add(key, value)
	End Sub
	
	Property Let CompareMode
	End Property
	
	Property Get Count
	End Property
	
	Function Exists(key)
	End Function
	
	Default Property Get Item(key)
	End Property
	
	Function Items(key)
	End Function
	
	Property Get Key(key)
	End Property
	
	Function Keys()
	End Property
	
	Sub Remove(key)
	End Sub
	
	Sub RemoveAll()
	End Sub
	
End Class

''' <summary>Contains information about run-time errors. Accepts the Raise and Clear methods for generating and clearing run-time errors.</summary>
Class Err

	''' <summary></summary>
	Sub Clear()
	End Sub


	''' <summary></summary>
	Property Get Description
	End Property


	''' <summary></summary>
	Property Get HelpContext
	End Property


	''' <summary></summary>
	Property Get HelpFile
	End Property


	''' <summary></summary>
	Property Get Number
	End Property


	 ''' <summary></summary>
	Sub Raise(number)
	End Sub
	
	 ''' <summary></summary>
	Sub Raise(number, source, description, helpfile, helpcontext)
	End Sub


	''' <summary></summary>
	Property Get Source
	End Property
	
End Class


''' <summary></summary>
Function Escape(str)
End Function


''' <summary></summary>
Function Eval(expr)
End Function


''' <summary></summary>
Function Exp(number)
End Function

Class FileSystemObject

	Function BuildPath(path, name)
	End Function

	Sub CopyFile(source, destination)
	End Sub
	Sub CopyFile(source, destination, overwrite)
	End Sub

	Sub CopyFolder(source, destination)
	End Sub
	Sub CopyFolder(source, destination, overwrite)
	End Sub

	Sub CreateFolder(foldername)
	End Sub

	Function CreateTextFile(filename)
	End Function
	Function CreateTextFile(filename, overwrite)
	End Function
	Function CreateTextFile(filename, overwrite, unicode)
	End Function

	Sub DeleteFile(filename)
	End Sub
	Sub DeleteFile(filename, force)
	End Sub

	Sub DeleteFolder(filename)
	End Sub
	Sub DeleteFolder(filename, force)
	End Sub

	Function DriveExists(drive)
	End Function

	Function FileExists(filename)
	End Function

	Function FolderExists(foldername)
	End Function

	Function GetAbsolutePathName(path)
	End Function

	Function GetBaseName(path)
	End Function

	Function GetDrive(drive)
	End Function

	Function GetDriveName(drive)
	End Function

	Function GetExtensionName(path)
	End Function

	Function GetFile(filename)
	End Function

	Function GetFileName(filename)
	End Function

	Function GetFolder(foldername)
	End Function

	Function GetParentFolderName(foldername)
	End Function

	Function GetSpecialFolder(folderspec)
	End Function

	Function GetTempName()
	End Function

	Sub MoveFile(source, destination)
	End Sub

	Sub MoveFolder(source, destination)
	End Sub

	Function OpenTextFile(filename)
	End Function
	Function OpenTextFile(filename, iomode)
	End Function
	Function OpenTextFile(filename, iomode, create)
	End Function
	Function OpenTextFile(filename, iomode, create, format)
	End Function

End Class


''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
Function Filter(InputStrings, Value)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
Function Filter(InputStrings, Value, Include)
End Function

''' <summary>Returns a zero-based array containing a subset of a string array based on a specified filter criteria.</summary>
Function Filter(InputStrings, Value, Include, Compare)
End Function


''' <summary></summary>
Function Fix(number)
End Function


''' <summary></summary>
Function FormatCurrency(Expression)
End Function

''' <summary></summary>
Function FormatCurrency(Expression, NumDigitsAfterDecimal)
End Function

''' <summary></summary>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit)
End Function

''' <summary></summary>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers)
End Function

''' <summary></summary>
Function FormatCurrency(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
End Function


''' <summary></summary>
Function FormatDateTime(Date)
End Function

''' <summary></summary>
Function FormatDateTime(Date, NamedFormat)
End Function


''' <summary></summary>
Function FormatNumber(Expression)
End Function

''' <summary></summary>
Function FormatNumber(Expression, NumDigitsAfterDecimal)
End Function

''' <summary></summary>
Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit)
End Function

''' <summary></summary>
Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers)
End Function

''' <summary></summary>
Function FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
End Function


''' <summary></summary>
Function FormatPercent(Expression)
End Function

''' <summary></summary>
Function FormatPercent(Expression, NumDigitsAfterDecimal)
End Function

''' <summary></summary>
Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit)
End Function

''' <summary></summary>
Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers)
End Function

''' <summary></summary>
Function FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
End Function


''' <summary></summary>
Property Get GetLocale
End Property


''' <summary>?</summary>
Function GetObject(pathname)
End Function

''' <summary>?</summary>
Function GetObject(pathname, classname)
End Function


''' <summary></summary>
Function GetRef(procname)
End Function


''' <summary></summary>
Property Get GetUILanguage
End Property


''' <summary></summary>
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
Function InStr(string1, string2)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStr(start, string1, string2)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
Function InStr(string1, string2, compare)
End Function

''' <summary>Returns the position of the first occurrence of one string within another.</summary>
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


''' <summary></summary>
Function Int(number)
End Function


''' <summary></summary>
Function IsArray(var)
End Function


''' <summary></summary>
Function IsDate(expr)
End Function


''' <summary></summary>
Function IsEmpty(expr)
End Function


''' <summary></summary>
Function IsNull(expr)
End Function


''' <summary>Returns a Boolean value indicating whether an expression can be evaluated as a number.</summary>
Function IsNumeric(expr)
End Function


''' <summary></summary>
Function IsObject(expr)
End Function


''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
Function Join(list)
End Function

''' <summary>Returns a string created by joining a number of substrings contained in an array.</summary>
Function Join(list, delimiter)
End Function


''' <summary></summary>
Function LBound(arrayname)
End Function

''' <summary></summary>
Function LBound(arrayname, dimension)
End Function


''' <summary></summary>
Function LCase(str)
End Function


''' <summary></summary>
Function Left(str, length)
End Function


''' <summary></summary>
Function LeftB(str, length)
End Function


''' <summary>Returns the number of characters in a string or the number of bytes required to store a variable.</summary>
Function Len(str)
End Function


''' <summary></summary>
Function LenB(str)
End Function


''' <summary></summary>
Function LoadPicture(picturename)
End Function


''' <summary></summary>
Function Log(number)
End Function


''' <summary></summary>
Function LTrim(str)
End Function

''' <summary>Provides access to the read-only properties of a regular expression match.</summary>
Class Match

	''' <summary></summary>
	Property Get FirstIndex
	End Property


	''' <summary></summary>
	Property Get Length
	End Property


	''' <summary></summary>
	Property Get Value
	End Property
End Class

''' <summary></summary>
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


''' <summary></summary>
Function Minute(time)
End Function


''' <summary></summary>
Function Month(date)
End Function


''' <summary></summary>
Function MonthName(date)
End Function

''' <summary></summary>
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


''' <summary>Returns the current date and time according to the setting of your computer's system date and time.)</summary>
Property Get Now
End Property


''' <summary></summary>
Function Oct(number)
End Function


''' <summary>Provides simple regular expression support</summary>
Class RegExp

	''' <summary></summary>
	Function Execute(str)
	End Function


	''' <summary></summary>
	Property Get Global
	End Property


	 ''' <summary></summary>
	Property Get IgnoreCase
	End Property


	 ''' <summary></summary>
	Property Get Pattern
	End Property
	

	''' <summary></summary>
	Function Replace(string1, string2)
	End Function


	''' <summary></summary>
	Function Test(str)
	End Function

End Class


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


''' <summary></summary>
Function Right(str, length)
End Function


''' <summary></summary>
Function RightB(str, length)
End Function


''' <summary></summary>
Function Rnd()
End Function

''' <summary></summary>
Function Rnd(number)
End Function


''' <summary></summary>
Function Round(number, digits)
End Function


''' <summary></summary>
Function RTrim(str)
End Function


''' <summary></summary>
Property Get ScriptEngine
End Property


''' <summary></summary>
Property Get ScriptEngineBuildVersion
End Property


''' <summary></summary>
Property Get ScriptEngineMajorVersion
End Property


''' <summary></summary>
Property Get ScriptEngineMinorVersion
End Property


''' <summary></summary>
Function Second(time)
End Function


''' <summary>undocumented</summary>
Function SetLocale(integer)
End Function


''' <summary></summary>
Function Sgn(number)
End Function


''' <summary></summary>
Function Sin(number)
End Function


''' <summary></summary>
Function Space(number)
End Function


''' <summary></summary>
Function Split(str)
End Function

''' <summary></summary>
Function Split(str, delimiter)
End Function

''' <summary></summary>
Function Split(str, delimiter, count)
End Function

''' <summary></summary>
Function Split(str, delimiter, count, compare)
End Function


''' <summary></summary>
Function Sqr(number)
End Function


''' <summary></summary>
Function StrComp(string1, string2)
End Function

''' <summary></summary>
Function StrComp(string1, string2, compare)
End Function


''' <summary></summary>
Function StrReverse(str)
End Function


''' <summary></summary>
Function Tan(number)
End Function



''' <summary></summary>
Property Get Time
End Property


''' <summary></summary>
Property Get Timer
End Property


''' <summary></summary>
Function TimeSerial(hour, minute, second)
End Function


''' <summary></summary>
Function TimeValue(time)
End Function


''' <summary></summary>
Function Trim(str)
End Function


''' <summary></summary>
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


''' <summary></summary>
Function Unescape(str)
End Function


''' <summary></summary>
Function VarType(var)
End Function


''' <summary></summary>
Function Weekday(date)
End Function

''' <summary></summary>
Function Weekday(date, firstdayofweek)
End Function


''' <summary></summary>
Function WeekdayName(weekday)
End Function

''' <summary></summary>
Function WeekdayName(weekday, abbreviate)
End Function

''' <summary></summary>
Function WeekdayName(weekday, abbreviate, firstdayofweek)
End Function

Class WScript

	''' <summary></summary>
	Function ConnectObject(objEventSource, strPrefix)
	End Function


	''' <summary></summary>
	Function CreateObject(strProgID)
	End Function

	''' <summary></summary>
	Function CreateObject(strProgID,strPrefix)
	End Function


	''' <summary></summary>
	Function DisconnectObject(obj)
	End Function


	''' <summary></summary>
	Function Echo(args)
	End Function


	''' <summary></summary>
	Function GetObject(strPathname)
	End Function

	''' <summary></summary>
	Function GetObject(strPathname, strProgID)
	End Function

	''' <summary></summary>
	Function GetObject(strPathname, strProgID, strPrefix)
	End Function


	''' <summary></summary>
	Function Quit()
	End Function

	''' <summary></summary>
	Function Quit(ErrorCode)
	End Function


	''' <summary></summary>
	Function Sleep(ms)
	End Function


	''' <summary></summary>
	Property Get Arguments
	End Property


	''' <summary></summary>
	Property Get BuildVersion
	End Property


	''' <summary></summary>
	Property Get FullName
	End Property


	''' <summary></summary>
	Property Get Interactive
	End Property


	''' <summary></summary>
	Property Get Name
	End Property


	''' <summary></summary>
	Property Get Path
	End Property


	''' <summary></summary>
	Property Get ScriptFullName
	End Property


	''' <summary></summary>
	Property Get ScriptName
	End Property


	''' <summary></summary>
	Property Get TimeOut
	End Property


	''' <summary></summary>
	Property Get Version
	End Property

End Class

''' <summary>Returns a whole number representing the year.</summary>
''' <param name="date">Any expression that can represent a date</param>
Function Year(date)
End Function


Const False = False ' Boolean
Const True = True
Const vbAbort = 3
Const vbAbortRetryIgnore = 2
Const vbApplicationModal = 0
Const vbArray = 8192
Const vbBinaryCompare = 0 ' Perform a binary comparison
Const vbBoolean = 11
Const vbByte = 17
Const vbCancel = 2
Const vbCr = Chr(13)
Const vbCritical = 16
Const vbCrLf = Chr(13) & Chr(10)
Const vbCurrency = 6
Const vbDatabaseCompare = 2
Const vbDataObject = 13
Const vbDate = 7
Const vbDecimal = 14
Const vbDefaultButton1 = 0
Const vbDefaultButton2 = 256
Const vbDefaultButton3 = 512
Const vbDefaultButton4 = 768
Const vbEmpty = 0
Const vbError = 10
Const vbExclamation = 18
Const vbFalse = 0
Const vbFirstFourDays = 2
Const vbFirstFullWeek = 3
Const vbFirstJan1 = 1
Const vbFormFeed = Chr(12)
Const vbFriday = 6
Const vbGeneralDate = 0
Const vbIgnore = 5
Const vbInformation =64
Const vbInteger = 2
Const vbLf = Chr(10)
Const vbLong = 3
Const vbLongDate = 1
Const vbLongTime = 3
Const vbMonday = 2
Const vbMsgBoxHelpButton = 16384
Const vbMsgBoxRight = 524288
Const vbMsgBoxRtlReading = 1048576
Const vbNewLine = Empty
Const vbNo = 7
Const vbNull = 1
Const vbNullChar = Chr(0)
Const vbNullString = Empty
Const vbObject = 9
Const vbObjectError = -2147221504
Const vbOK = 1
Const vbOKCancel = 1 ' Display OK and Cancel buttons
Const vbOKOnly = 0 ' Display OK button only.
Const vbQuestion = 32 ' Display Warning Query icon.
Const vbRetry = 4 ' Retry button was clicked
Const vbRetryCancel = 5 ' Display Retry and Cancel buttons.
Const vbSaturday = 7
Const vbShortDate = 2
Const vbShortTime = 4
Const vbSingle = 4
Const vbString = 8
Const vbSunday = 1
Const vbSystemModal = 4096
Const vbTab = Chr(9)
Const vbTextCompare = 1 ' Perform a textual comparison
Const vbThursday = 5
Const vbTrue = -1
Const vbTuesday = 3
Const vbUseDefault = -2
Const vbUseSystem = 0
Const vbUseSystemDayOfWeek = 0
Const vbVariant = 12
Const vbVerticalTab = Chr(11)
Const vbWednesday = 4
Const vbYes = 6
Const vbYesNo = 4
Const vbYesNoCancel = 3

Const vbBlack = &h00
Const vbBlue = &hFF0000
Const vbCyan = &hFFFF00
Const vbGreen = &hFF00
Const vbMagenta = &hFF00FF
Const vbRed = &hFF
Const vbWhite = &hFFFFFF
Const vbYellow = &hFFFF
