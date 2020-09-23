
class Klasse1

    private m_Text

    Public Property Get Text
        Text = m_Text
    end Property
    Public Property Let Text(value)
        m_Text = value
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

End Class