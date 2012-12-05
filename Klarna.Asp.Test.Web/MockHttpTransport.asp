<%

'------------------------------------------------------------------------------
' Mocked HTTP Transport.
'------------------------------------------------------------------------------
Class MockHttpTransport
    Public m_url

    Public Sub CreateRequest(url)
        m_url = url
    End Sub
End Class

%>
