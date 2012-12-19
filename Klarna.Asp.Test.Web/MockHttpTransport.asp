<%

'------------------------------------------------------------------------------
' Mocked HTTP Transport.
'------------------------------------------------------------------------------
Class MockHttpTransport
    Public m_request
    Public m_requestInSend
    Public m_response

    Public Function CreateRequest(url)
        m_request.setUri(url)
        Set CreateRequest = m_request
    End Function

    Public Function Send(request)
        Set m_requestInSend = request
        Set Send = m_response
    End Function

End Class

%>
