<%

'------------------------------------------------------------------------------
' Mocked connector.
'------------------------------------------------------------------------------
Class MockConnector
    Public m_httpMethod
    Public m_resource
    Public m_options

    Public Function Apply(ByVal httpMethod, ByVal resource, ByVal options)
        m_httpMethod = httpMethod
        Set m_resource = resource
        Set m_options = options
    End Function

    Public Function GetBaseUri()
        GetBaseUri = "http://stub.com"
    End Function

End Class

%>
