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
' The HTTP request.
'------------------------------------------------------------------------------

Class HttpRequest
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_headers
    Private m_uri
    Private m_method
    Private m_data

    ' -------------------------------------------------------------------------
    ' Class constructor
    '
    ' Initializes a new instance of the Request class.
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_headers = Server.CreateObject("Scripting.Dictionary")
        m_method = "GET"
        m_data = ""
    End Sub

    Private Sub Class_Terminate
        Set m_headers = Nothing
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the request uri.
    ' -------------------------------------------------------------------------
    Public Function GetUri()
        GetUri = m_uri
    End Function

    Public Function SetUri(uri)
        m_uri = uri
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the HTTP method used for the request.
    ' -------------------------------------------------------------------------
    Public Function GetMethod()
        GetMethod = m_method
    End Function

    Public Function SetMethod(method)
        m_method = UCase(method)
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets headers.
    ' -------------------------------------------------------------------------
    Public Sub SetHeader(name, value)
        If m_headers.Exists(name) Then
            m_headers.Item(name) = value
        Else
            m_headers.Add name, value
        End If
    End Sub

    Public Function GetHeader(name)
        If m_headers.Exists(name) Then
            GetHeader = m_headers.Item(name)
        Else
            GetHeader = ""
        End If
    End Function

    Public Function GetHeaders()
        Set GetHeaders = m_headers
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the data (payload) for the request.
    ' -------------------------------------------------------------------------
    Public Function GetData()
        GetData = m_data
    End Function

    Public Function SetData(data)
        m_data = data
    End Function

End Class

%>