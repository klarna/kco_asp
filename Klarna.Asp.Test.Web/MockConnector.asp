<%

'------------------------------------------------------------------------------
' Mocked connector.
'------------------------------------------------------------------------------
Class MockConnector
    Public m_httpMethod
    Public m_order
    Public m_options

    Public Function Apply(ByVal httpMethod, ByVal order, ByVal options)
        m_httpMethod = httpMethod
        Set m_order = order
        Set m_options = options
    End Function
End Class

%>
