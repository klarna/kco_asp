<%
'------------------------------------------------------------------------------
'   Copyright 2012 Klarna AB
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'       http://www.apache.org/licenses/LICENSE-2.0
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
' 
'   Klarna Support: support@klarna.com
'   http://integration.klarna.com/
'------------------------------------------------------------------------------

'--------------------------------------------------------------------------
' Creates a new instance or the BasicConnector class.
'
' Parameters:
' object    httpTransport    The HTTP transport to use.
' object    digest           The digest to use.
' string    secret           The secret to use.
'
' Returns:
' The basic connector instance.
'--------------------------------------------------------------------------
Public Function CreateBasicConnector(httpTransport, digest, secret)
    Dim connector
    Set connector = new BasiConnector
    connector.SetHttpTransport(httpTransport)
    connector.SetDigest(digest)
    connector.SetSecret(secret)

    Set CreateBasicConnector = connector
End Function

'------------------------------------------------------------------------------
' The basic connector.
'------------------------------------------------------------------------------
Class BasicConnector
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_userAgent
    Private m_httpTransport
    Private m_digest
    Private m_secret

    ' -------------------------------------------------------------------------
    ' Class constructor
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_userAgent = new UserAgent
    End Sub

    Private Sub Class_Terminate
    End Sub

    Public Sub SetHttpTransport(httpTransport)
        Set m_httpTransport = httpTransport
    End Sub

    Public Sub SetDigest(digest)
        Set m_digest = digest
    End Sub

    Public Sub SetSecret(secret)
        m_secret = secret
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the user agent used for User-Agent header.
    ' -------------------------------------------------------------------------
    Public Function GetUserAgent()
        Set GetUserAgent = m_userAgent
    End Function

    Public Function SetUserAgent(ua)
        Set m_userAgent = ua
    End Function

    ' -------------------------------------------------------------------------
    ' Applies a HTTP method on a specific resource.
    ' -------------------------------------------------------------------------
    Public Function Apply(httpMethod, order, options)
        ' Call Handle(...)
    End Function

End Class

%>