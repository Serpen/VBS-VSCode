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

	Property Get Attributes
	End Property

	Function Copy(Destination)
	End Function
	Function Copy(Destination, OverWriteFiles)
	End Function

	Property Get DateCreated
	End Property

	Property Get DateLastAccessed
	End Property
	
	Property Get DateLastModified
	End Property

	Function Delete()
	End Function
	Function Delete(Force)
	End Function

	Property Get Drive
	End Property

	Function Move(Destination)
	End Function

	Property Get Name
	End Property

	Function OpenAsTextStream()
	End Function
	Function OpenAsTextStream(IOMode)
	End Function
	Function OpenAsTextStream(IOMode, Format)
	End Function

	Property Get ParentFolder
	End Property

	Property Get Path
	End Property

	Property Get ShortName
	End Property

	Property Get ShortPath
	End Property

	Property Get Size
	End Property

	Property Get Type
	End Property

End Class

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

	Function CreateFolder(foldername)
	End Function

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

	Function GetFileVersion(filename)
	End Function

	Function GetFolder(foldername)
	End Function

	Function GetParentFolderName(foldername)
	End Function

	Function GetSpecialFolder(folderspec)
	End Function
	
	Function GetStandardStream(StandardStreamType)
	End Function
	Function GetStandardStream(StandardStreamType, Unicode)
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

''' <summary>Provides access to the read-only properties of a regular expression match.</summary>
Class Match

	Property Get FirstIndex
	End Property

	Property Get Length
	End Property

	Property Get Value
	End Property
End Class

Class RegExp

	Function Execute(str)
	End Function

	Property Get Global
	End Property

	Property Get IgnoreCase
	End Property

	Property Get Pattern
	End Property
	

	Function Replace(string1, string2)
	End Function

	Function Test(str)
	End Function

End Class

Class TextStream

	Property Get AtEndOfLine
	End Property

	Property Get AtEndOfStream
	End Property

	Sub Close()
	End Sub

	Property Get Column
	End Property

	Property Get Line
	End Property

	Function Close()
	End Function

	Function Read(chars)
	End Function

	Function ReadAll()
	End Function

	Function ReadLine()
	End Function

	Sub Skip(chars)
	End Sub

	Sub SkipLine(chars)
	End Sub

	Sub Write(text)
	End Sub

	Sub WriteBlankLines(lines)
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

	Property Get BuildVersion
	End Property
	
	Function ConnectObject(objEventSource, strPrefix)
	End Function

	Function CreateObject(strProgID)
	End Function

	Function CreateObject(strProgID,strPrefix)
	End Function

	Function DisconnectObject(obj)
	End Function

	Function Echo(args)
	End Function

	Property Get FullName
	End Property

	Function GetObject(strPathname)
	End Function

	Function GetObject(strPathname, strProgID)
	End Function

	Function GetObject(strPathname, strProgID, strPrefix)
	End Function

	Property Get Interactive
	End Property

	Property Get Name
	End Property

	Property Get Path
	End Property

	Function Quit()
	End Function

	Function Quit(ErrorCode)
	End Function

	Property Get ScriptFullName
	End Property

	Property Get ScriptName
	End Property

	Function Sleep(ms)
	End Function

	Property Get StdErr
	End Property

	Property Get StdIn
	End Property

	Property Get StdOut
	End Property

	Property Get TimeOut
	End Property

	Property Get Version
	End Property

End Class
