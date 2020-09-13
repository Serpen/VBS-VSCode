Option Explicit
dim retValue

dim fso : set fso = CreateObject("Scripting.FileSystemObject")

dim file : file = "C:\temp\wscript.txt"

On Error Resume Next
if UBound(Wscript.Arguments) > 0 Then
    file = Wscript.Arguments(0)
Else
    file = ".\TestFile.vbs"
end if
On Error Goto 0

dim stream : set stream = fso.OpenTextFile(file)

Do Until stream.AtEndOfStream
    dim line : line = stream.ReadLine()
    dim done : done = false
    retValue = ""
    On Error Resume Next
    ExecuteGlobal "Option Explicit: retValue = " & line
    Select Case Err.Number
        Case 438
            done = true
        case 500
            done = true
        Case 0
            Wscript.echo  line & " : " & Err.Description & " " & Err.Number & " " & retValue
            done = true
        Case Else
            Wscript.echo  line & " : " & Err.Description & " " & Err.Number & " " & retValue
            done = true
    End Select

    If not Done Then
        On Error Resume Next
        ExecuteGlobal "retValue = " & line & ".SurelyNotThere"
        Select Case Err.Number
            ' Case 438
                ' done = true
            case 500
                done = true
            Case 0
                Wscript.echo  line & " : " & Err.Description & " " & Err.Number & " " & retValue
                done = true
            Case Else
                Wscript.echo  line & " : " & Err.Description & " " & Err.Number & " " & retValue
                done = true
        End Select
    End If
Loop

stream.close