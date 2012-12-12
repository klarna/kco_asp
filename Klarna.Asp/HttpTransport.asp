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
        m_timeout = 5000
    End Sub

    Private Sub Class_Terminate
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the number of milliseconds before the connection times out.
    ' -------------------------------------------------------------------------
    Public Function GetTimeout()
        GetTimeout = m_timeout
    End Function

    Public Function SetTimeout(timeout)
        m_timeout = timeout
    End Function

    ' -------------------------------------------------------------------------
    ' Creates a HttpRequest.
    ' -------------------------------------------------------------------------
    Public Function CreateRequest(uri)
        Dim hr
        Set hr = New HttpRequest
        hr.SetUri(uri)

        Set CreateRequest = hr
    End Function

End Class
%>