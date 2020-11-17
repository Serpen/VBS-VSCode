Option Explicit

''' <summary>An intrinsic global object that can send output to a script debugger, such as the Microsoft Script Debugger.</summary>
Class Debug

	Sub Write(str)
	End Sub
	
	
	Sub WriteLine(str)
	End Sub
End Class


''' <summary>Object that stores data key, item pairs.</summary>
Class Dictionary
	
	Sub Add(key, value)
	End Sub
	
	Property Get CompareMode ' As Long
	End Property
	Property Let CompareMode(Mode)
	End Property
	
	Property Get Count ' As Long
	End Property
	
	Function Exists(key) ' as Boolean
	End Function
	
	Property Get HashVal(key) ' As Long
	End Property
	
	Public Default Property Get Item(key)
	End Property
	
	Function Items(key) ' As Variant
	End Function
	
	Property Get Key(key)
	End Property
	
	Function Keys() ' As Variant
	End Function
	
	Sub Remove(key)
	End Sub

	Sub RemoveAll()
	End Sub
	
End Class


''' <summary>Contains information about run-time errors. Accepts the Raise and Clear methods for generating and clearing run-time errors.</summary>
Class Err

	Sub Clear()
	End Sub

	Property Get Description
	End Property

	Property Get HelpContext
	End Property

	Property Get HelpFile
	End Property

	Property Get Number
	End Property

	Sub Raise(number)
	End Sub
	
	Sub Raise(number, source, description, helpfile, helpcontext)
	End Sub

	Property Get Source
	End Property
	
End Class


Class File

	Property Get Attributes ' as Long
	End Property

	Sub Copy(Destination)
	End Sub
	Sub Copy(Destination, OverWriteFiles)
	End Sub

	Property Get DateCreated ' as Date
	End Property

	Property Get DateLastAccessed ' as Date
	End Property
	
	Property Get DateLastModified ' as Date
	End Property

	Sub Delete()
	End Sub
	Sub Delete(Force)
	End Sub

	Property Get Drive ' as Drive
	End Property

	Sub Move(Destination)
	End Sub

	Property Get Name ' As String
	End Property

	Function OpenAsTextStream() ' As TextStream
	End Function
	Function OpenAsTextStream(IOMode) ' As TextStream
	End Function
	Function OpenAsTextStream(IOMode, Format) ' As TextStream
	End Function

	Property Get ParentFolder ' As Folder
	End Property

	Property Get Path ' As String
	End Property

	Property Get ShortName ' As String
	End Property

	Property Get ShortPath ' As String
	End Property

	Property Get Size ' as Long
	End Property

	Property Get Type ' As String
	End Property

End Class


Class Folder

	Property Get Attributes ' as Long
	End Property

	Sub Copy(Destination)
	End Sub
	Sub Copy(Destination, OverWriteFiles)
	End Sub

	Property Get DateCreated ' as Date
	End Property

	Property Get DateLastAccessed ' as Date
	End Property
	
	Property Get DateLastModified ' as Date
	End Property

	Sub Delete()
	End Sub
	Sub Delete(Force)
	End Sub

	Property Get Drive ' as Drive
	End Property

	Property Get Files ' as FileCollection
	End Property

	Property Get IsRootFolder ' as Boolean
	End Property

	Sub Move(Destination)
	End Sub

	Property Get Name ' As String
	End Property

	Function CreateTextFile(FileName) ' As TextStream
	End Function
	Function CreateTextFile(FileName, Overwrite) ' As TextStream
	End Function
	Function CreateTextFile(FileName, Overwrite, Unicode) ' As TextStream
	End Function

	Property Get ParentFolder ' As Folder
	End Property

	Property Get Path ' As String
	End Property

	Property Get ShortName ' As String
	End Property

	Property Get ShortPath ' As String
	End Property

	Property Get Size ' as Long
	End Property

	Property Get SubFolders ' as FolderCollection
	End Property

	Property Get Type ' As String
	End Property

End Class


Class FileSystemObject

	Function BuildPath(path, name) ' As String
	End Function

	Sub CopyFile(source, destination)
	End Sub
	Sub CopyFile(source, destination, overwrite)
	End Sub

	Sub CopyFolder(source, destination)
	End Sub
	Sub CopyFolder(source, destination, overwrite)
	End Sub

	Function CreateFolder(foldername) ' As Folder
	End Function

	Function CreateTextFile(filename) ' As TextStream
	End Function
	Function CreateTextFile(filename, overwrite) ' As TextStream
	End Function
	Function CreateTextFile(filename, overwrite, unicode) ' As TextStream
	End Function

	Sub DeleteFile(filename)
	End Sub
	Sub DeleteFile(filename, force)
	End Sub

	Sub DeleteFolder(filename)
	End Sub
	Sub DeleteFolder(filename, force)
	End Sub

	Property Get Drives ' As DriveCollection
	End Property

	Function DriveExists(drive) ' As Boolean
	End Function

	Function FileExists(filename) ' As Boolean
	End Function

	Function FolderExists(foldername) ' As Boolean
	End Function

	Function GetAbsolutePathName(path) ' As String
	End Function

	Function GetBaseName(path) ' As String
	End Function

	Function GetDrive(drive) ' As Drive
	End Function

	Function GetDriveName(drive) ' As String
	End Function

	Function GetExtensionName(path) ' As String
	End Function

	Function GetFile(filename) ' As File
	End Function

	Function GetFileName(filename) ' As String
	End Function

	Function GetFileVersion(filename) ' As String
	End Function

	Function GetFolder(foldername) ' As Folder
	End Function

	Function GetParentFolderName(foldername) ' As String
	End Function

	Function GetSpecialFolder(folderspec) ' As Folder
	End Function
	
	Function GetStandardStream(StandardStreamType) ' As TextStream
	End Function
	Function GetStandardStream(StandardStreamType, Unicode) ' As TextStream
	End Function

	Function GetTempName() ' As String
	End Function

	Sub MoveFile(source, destination)
	End Sub

	Sub MoveFolder(source, destination)
	End Sub

	Function OpenTextFile(filename) ' As TextStream
	End Function
	Function OpenTextFile(filename, iomode) ' As TextStream
	End Function
	Function OpenTextFile(filename, iomode, create) ' As TextStream
	End Function
	Function OpenTextFile(filename, iomode, create, format) ' As TextStream
	End Function
	
End Class


Class Drive
	Property Get AvailableSpace ' As Double
	End Property

	Property Get DriveLetter ' As String
	End Property

	Property Get DriveType ' As Long
	End Property

	Property Get FileSystem ' As String
	End Property

	Property Get FreeSpace ' As Double
	End Property

	Property Get IsReady ' As Boolean
	End Property

	Property Get Path ' As String
	End Property

	Property Get RootFolder ' As Folder
	End Property

	Property Get SerialNumber ' As Long
	End Property

	Property Get ShareName ' As String
	End Property

	Property Get TotalSize ' As Double
	End Property

	Property Get VolumeName ' As String
	End Property

End Class


''' <summary>Provides access to the read-only properties of a regular expression match.</summary>
Class Match

	Property Get FirstIndex ' As Long
	End Property

	Property Get Length ' As Long
	End Property

	Property Get SubMatches ' As String()
	End Property

	Property Get Value ' As String
	End Property

End Class


Class RegExp

	Function Execute(str) ' as Object
	End Function

	Property Get Global ' As Boolean
	End Property
	Property Let Global(b)
	End Property

	Property Get IgnoreCase ' As Boolean
	End Property
	Property Let IgnoreCase(b)
	End Property

	Property Get Pattern ' As String
	End Property
	Property Let Pattern(s)
	End Property
	
	Function Replace(string1, string2) ' As String
	End Function

	Function Test(str) ' As Boolean
	End Function

End Class


Class TextStream

	Property Get AtEndOfLine ' As Boolean
	End Property

	Property Get AtEndOfStream ' As Boolean
	End Property

	Sub Close()
	End Sub

	Property Get Column ' As Long
	End Property

	Property Get Line ' As Long
	End Property

	Function Read(Characters) ' As String
	End Function

	Function ReadAll() ' As String
	End Function

	Function ReadLine() ' As String
	End Function

	Sub Skip(Characters)
	End Sub

	Sub SkipLine()
	End Sub

	Sub Write(Text)
	End Sub

	Sub WriteBlankLines(Lines)
	End Sub

	Sub WriteLine()
	End Sub
	Sub WriteLine(text)
	End Sub

End Class


Class WScript

	Property Get Application
	End Property

	Property Get Arguments
	End Property

	Property Get BuildVersion ' As String
	End Property
	
	Sub ConnectObject(objEventSource, strPrefix)
	End Sub

	Function CreateObject(strProgID) ' As Object
	End Function

	Function CreateObject(strProgID, strPrefix) ' As Object
	End Function

	Sub DisconnectObject(obj)
	End Sub

	Sub Echo(args)
	End Sub

	Property Get FullName ' As String
	End Property

	Function GetObject(strPathname) ' As Object
	End Function

	Function GetObject(strPathname, strProgID) ' As Object
	End Function

	Function GetObject(strPathname, strProgID, strPrefix) ' As Object
	End Function

	Property Get Interactive ' As Boolean
	End Property

	Property Get Name ' As String
	End Property

	Property Get Path ' As String
	End Property

	Sub Quit()
	End Sub

	Sub Quit(ErrorCode)
	End Sub

	Property Get ScriptFullName ' As String
	End Property

	Property Get ScriptName ' As String
	End Property

	Sub Sleep(ms)
	End Sub

	Property Get StdErr ' As TextStream
	End Property

	Property Get StdIn ' As TextStream
	End Property

	Property Get StdOut ' As TextStream
	End Property

	Property Get TimeOut ' As Integer
	End Property

	Property Get Version ' As String
	End Property

End Class

Class Shell

	Function AppActivate(App, Wait) ' As Boolean
	End Function

	Function CreateShortcut(PathLink)
	End Function
	
	Function Exec(Command)
	End Function
	
	Function ExpandEnvironmentStrings(Src) ' As String
	End Function
	
	Function LogEvent(Type, Message, Target) ' As Boolean
	End Function
	
	Function Popup(Text, SecondsToWait, Title, Type) ' As Integer
	End Function
	
	Sub RegDelete(Name)
	End Sub
	
	Function RegRead(Name)
	End Function
	
	Sub RegWrite(Name, Value, Type)
	End Sub
	
	Function Run(Command, WindowStyle, WaitOnReturn) ' As Integer
	End Function
	
	Sub SendKeys(Keys, Wait)
	End Sub

	Public Default Property Get Environment(Type)
	End Property

	Property Get CurrentDirectory
	End Property
	
	Property Let CurrentDirectory
	End Property

	Property Get SpecialFolders
	End Property

End Class

Private Class Picture

	Property Get Handle ' As Long
	End Property

	Sub Render(hdc, x, y, cx, cy, xSrc, ySrc, cxSrc, cySrc, prcWBounds)
	End Sub

	Property Get Height ' As Long
	End Property

	Property Get hPal ' As Long
	End Property

	Property Get Type ' As Integer
	End Property

	Property Get Width ' As Long
	End Property

End Class