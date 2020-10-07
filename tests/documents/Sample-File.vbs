
' Execute a SELECT query against database 
Function ExecuteQuery(sql)
    Dim querystring : querystring = Trim(sql)

    dim obj
    set obj = new Data

End Function

Class DataBaseQuery

    Private m_query
    Property Get Query
        Query = m_query
    End Property

    Property Let Query(val)
        m_query = val
    End Property

    ' Real work
    Public Sub Execute()
        ' Todo: Implement!
    End Sub
End Class