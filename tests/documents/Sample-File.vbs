

' Execute a SELECT query against database 
Function ExecuteQuery(sql)
    Dim querystring : querystring = Trim(sql)

    dim obj
    set obj = new DataBase
    obj.Execute()

End Function

' Class for Database access
Class DataBase

    Private m_query
    Property Get Query
        Query = m_query
    End Property

    Property Let Query(val)
        m_query = val
    End Property

    ''' <summary>Executes against a DB</summary>
    ''' <param name="force">Ignore Errors</param>
    Public Sub Execute(force)
        ' Todo: Implement!
    End Sub
End Class