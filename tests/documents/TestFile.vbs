Option Explicit

Main

Sub Main()

    dim fso : set fso = CreateObject("Scripting.FileSystemObject")

End Sub

Const myConst = 4
Dim firstVar
dim secondVar : dim Thirdwar


Public sub KlassenConstruct()
    dim cls
    
    set cls = new Klasse1
    With cls
        msgbox .Text
    End With
end sub

If True Then
    
End If

if firstVar = "" then Abs(8)

For secondVar=0 To myConst Step 1

Next

While false
   
WEnd

Class Klasse1

    public Property Get Text
        Text = "Test"
    end Property

    Private Sub Class_Initialize()
        ' Called automatically when class is created
    End Sub

    Private Sub Class_Terminate()
        ' Called automatically when all references to class instance are removed
    End Sub


End Class
