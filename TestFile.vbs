Option Explicit


Main

Sub Main()

dim fso : set fso = CreateObject("Scripting.FileSystemObject")
dim testVar4

MsgBox Wscript.ReadLine

WScript.Quit(1)

Wscript.StdErr
Wscript.StdIn
Wscript.StdOut

End Sub


Const myConst = 4
Dim firstVar
dim secondVar : dim Thirdwar
dim WorldWar
Dim EndVar

Function Funktion(test1, test2)
    Call Subroutine(test1)
    'msgbox myConst    
End Function

private sub Subroutine(first)
    Call Abs(myConst)
End Sub

Public sub KlassenConstruct()
    dim cls
    
    set cls = new Klasse1
    With cls
        'msgbox .GetText
    End With
end sub

If True Then
    
End If

if firstVar = "" then Abs(8)

For secondVar=0 To myConst Step 1

Next

While false
   
WEnd

Call Funktion(1,2)

Call Subroutine(1)

KlassenConstruct

class Klasse1

    public Property Get GetText
        GetText = "Test"
    end Property

end Class

Class Klasse2

    Private Sub Class_Initialize()
        ' Called automatically when class is created
    End Sub

    Private Sub Class_Terminate()
        ' Called automatically when all references to class instance are removed
    End Sub


End Class

Class Klasse3

    Private Sub Class_Initialize()
        ' Called automatically when class is created
    End Sub

    Private Sub Class_Terminate()
        ' Called automatically when all references to class instance are removed
    End Sub


End Class