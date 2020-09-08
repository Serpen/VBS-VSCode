Const myConst = 4

dim myvar

Function Funktion(test1, test2)
    Call Subroutine(test1)
    msgbox myConst
End Function

Private Sub Subroutine(first)
    Call Abs(myConst)
    
End Sub

Public sub KlassenConstruct()
    dim cls
    set cls = new Klasse
    With cls
        msgbox .GetText
    End With
end sub

If True Then
    
End If

if Hallo = "" then Abs(8)

For i=0 To Value Step 1
    
Next

While false
    
WEnd

Call Funktion(1,2)

Call Subroutine(1)

KlassenConstruct

class Klasse

    public Property Get GetText
        GetText = "Test"
    end Property
end Class