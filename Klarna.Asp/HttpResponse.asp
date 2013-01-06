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
' The HTTP response.
'------------------------------------------------------------------------------

Class HttpResponse
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_headers
    Private m_status
    Private m_data

    ' -------------------------------------------------------------------------
    ' Class constructor
    '
    ' Initializes a new instance of the Request class.
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_headers = Server.CreateObject("Scripting.Dictionary")
        m_status = 0
        m_data = ""
    End Sub

    Private Sub Class_Terminate
        Set m_headers = Nothing
    End Sub

    ' -------------------------------------------------------------------------
    ' Creates the HttpResponse.
    '
    ' Parameters:
    ' int       status          The status code.
    ' string    headers         The headers.
    ' string    data            The data.
    ' -------------------------------------------------------------------------
    Public Sub Create(status, headers, data)
        m_status = CInt(status)
        ParseHeaders headers
        m_data = data
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets the HTTP status code.
    '
    ' Returns:
    ' int   The status code.
    ' -------------------------------------------------------------------------
    Public Function GetStatus()
        GetStatus = m_status
    End Function

    ' -------------------------------------------------------------------------
    ' Gets the headers for the response.
    '
    ' Parameter:
    ' string    name    The header name.
    '
    ' Returns:
    ' string    The header value.
    ' -------------------------------------------------------------------------

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
    ' Gets the data (payload) for the response.
    '
    ' Returns:
    ' string    The data.
    ' -------------------------------------------------------------------------
    Public Function GetData()
        GetData = m_data
    End Function

    '--------------------------------------------------------------------------
    ' Parses headers from one string.
    ' Assumes that each header is separated by vbCrLf and name / value pairs
    ' are separated by first occurence of :
    '
    ' Parameter:
    ' string    headerString    The headers.
    '--------------------------------------------------------------------------
    Private Sub ParseHeaders(headerString)
        m_headers.RemoveAll

        Dim headers
        headers = Split(headerString, vbCrLf)
        Dim header
        For Each header in headers
            Dim splitPosition
            splitPosition = InStr(header, ":")
            If Len(header) > 2 And splitPosition > 0 Then
                m_headers.Add Trim(Left(header, splitPosition-1)), _
                    Trim(Right(header, Len(header) - splitPosition))
            End If
        Next
    End Sub
End Class

%>