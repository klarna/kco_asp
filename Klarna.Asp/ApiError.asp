<!-- #include file="json2.asp" -->

<%
'------------------------------------------------------------------------------
'   Copyright 2015 Klarna AB
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'
'   Klarna Support: support@klarna.com
'   http://developers.klarna.com/
'------------------------------------------------------------------------------

'--------------------------------------------------------------------------
' Creates a new instance of the ApiError class.
'
' Parameters:
' object    response     The HttpResponse object.
'
' Returns:
' The ApiError instance
'--------------------------------------------------------------------------
Public Function CreateApiError(response)
    Dim apiErr, data

    On Error Resume Next

    Set apiErr = new ApiError

    apiErr.SetResponse(response)

    data = response.GetData()

    If Len(data) > 0 Then
        apiErr.Parse(data)
    End If

    If Err.Number <> 0 Then
        ' Error occurred during parsing of the error response, ignore it.

        Err.Clear
    End If

    Set CreateApiError = apiErr
End Function

'------------------------------------------------------------------------------
' The API error resource.
'------------------------------------------------------------------------------
Class ApiError
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_resourceData
    Private m_response

    ' -------------------------------------------------------------------------
    ' Class constructor
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_resourceData = Nothing
        Set m_response = Nothing
    End Sub

    Private Sub Class_Terminate
        Set m_resourceData = Nothing
        Set m_response = Nothing
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the underlying HTTP response object for this error.
    '
    ' Parameter/Returns:
    ' object    The HttpResponse object.
    ' -------------------------------------------------------------------------
    Public Function GetResponse()
        Set GetResponse = m_response
    End Function

    Public Function SetResponse(response)
        Set m_response = response
    End Function

    ' -------------------------------------------------------------------------
    ' Replace resource with the new data.
    '
    ' Parameter:
    ' string    The data.
    ' -------------------------------------------------------------------------
    Public Function Parse(data)
        Set m_resourceData = JSON.parse(data)
    End Function

    ' -------------------------------------------------------------------------
    ' Basic representation of the resource.
    ' -------------------------------------------------------------------------
    Public Function Marshal()
        Set Marshal = m_resourceData
    End Function

    ' -------------------------------------------------------------------------
    ' Basic representation of the resource, in JSON format.
    ' -------------------------------------------------------------------------
    Public Function MarshalAsJson()
        MarshalAsJson = JSON.stringify(m_resourceData)
    End Function

End Class

%>
