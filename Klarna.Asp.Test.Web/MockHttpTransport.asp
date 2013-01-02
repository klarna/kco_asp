<%

'------------------------------------------------------------------------------
' Mocked HTTP Transport.
'------------------------------------------------------------------------------
Class MockHttpTransport
    Public m_request
    Public m_requestInSend
    Public m_response
    Public m_response2
    Public m_responseCount

    Public Function CreateRequest(url)
        m_request.setUri(url)
        Set CreateRequest = m_request
    End Function

    Public Function Send(request)
        Set m_requestInSend = request
        m_responseCount = m_responseCount + 1

        If m_responseCount < 2 Then
            Set Send = m_response
        Else
            Set Send = m_response2
        End If
    End Function

End Class

%>
