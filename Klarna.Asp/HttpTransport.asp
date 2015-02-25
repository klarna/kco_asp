<%
'------------------------------------------------------------------------------
'   Copyright 2013 Klarna AB
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
'   http://developers.klarna.com/
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' The HTTP transport.
'------------------------------------------------------------------------------

Class HttpTransport
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_timeout

    ' -------------------------------------------------------------------------
    ' Class constructor
    '
    ' Initializes a new instance of the HttpTransport class.
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        m_timeout = 10000
    End Sub

    Private Sub Class_Terminate
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the number of milliseconds before the connection times out.
    '
    ' Parameter/Returns:
    ' int    The timeout in milliseconds.
    ' -------------------------------------------------------------------------
    Public Function GetTimeout()
        GetTimeout = m_timeout
    End Function

    Public Function SetTimeout(timeout)
        m_timeout = timeout
    End Function

    ' -------------------------------------------------------------------------
    ' Creates a HttpRequest.
    '
    ' Parameters:
    ' string    uri             The uri.
    '
    ' Returns:
    ' object    The request.
    ' -------------------------------------------------------------------------
    Public Function CreateRequest(uri)
        Dim hr
        Set hr = New HttpRequest
        hr.SetUri(uri)

        Set CreateRequest = hr
    End Function

    ' -------------------------------------------------------------------------
    ' Performs a HTTP request.
    '
    ' Parameters:
    ' object    httpRequest     The request.
    '
    ' Returns:
    ' object    The response.
    ' -------------------------------------------------------------------------
    Public Function Send(httpRequest)
        Dim xmlHttp
        Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

        const RESOLVE_TIMEOUT = 30000
        const SEND_TIMEOUT = 30000
        const RECEIVE_TIMEOUT =30000
        xmlHttp.setTimeouts RESOLVE_TIMEOUT, m_timeout, SEND_TIMEOUT, RECEIVE_TIMEOUT

        xmlhttp.open httpRequest.GetMethod, httpRequest.GetUri, false

        Dim headers
        Set headers = httpRequest.GetHeaders
        Dim key
        Dim requestHeaders
        For Each key in headers.Keys
            xmlHttp.setRequestHeader key, headers.Item(key)
        Next

        If httpRequest.GetMethod = "POST" Then
            xmlhttp.send(httpRequest.GetData)
        Else
            xmlhttp.send()
        End If

        Dim result
        Set result = New HttpResponse
        result.Create xmlHttp.status, xmlHttp.getAllResponseHeaders, xmlHttp.responseText

        Set Send = result
    End Function

End Class
%>