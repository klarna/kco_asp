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
' Creates a new instance of the RecurringStatus class.
'
' Parameters:
' object    connector    The connector to use.
' string    token        The recurring token.
'
' Returns:
' The recurring status instance.
'--------------------------------------------------------------------------
Public Function CreateRecurringStatus(connector, token)
    Dim status
    Set status = new RecurringStatus

    status.SetConnector(connector)
    status.SetToken(token)
    status.SetContentType "application/vnd.klarna.checkout.recurring-status-v1+json"

    Set CreateRecurringStatus = status
End Function

'------------------------------------------------------------------------------
' The recurring status resource.
'------------------------------------------------------------------------------
Class RecurringStatus
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_resourceData
    Private m_connector
    Private m_location
    Private m_accept
    Private m_contentType
    Private m_path
    Private m_error

    ' -------------------------------------------------------------------------
    ' Class constructor
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_resourceData = Nothing
        Set m_connector = Nothing
        Set m_error = Nothing
        m_location = ""
        m_accept = ""
        m_contentType = ""
        m_path = ""
    End Sub

    Private Sub Class_Terminate
        Set m_resourceData = Nothing
        Set m_connector = Nothing
        Set m_error = Nothing
    End Sub

    ' -------------------------------------------------------------------------
    ' Sets connector to use.
    '
    ' Parameter:
    ' object    The connector.
    ' -------------------------------------------------------------------------
    Public Sub SetConnector(connector)
        Set m_connector = connector
    End Sub

    ' -------------------------------------------------------------------------
    ' Sets the recurring token.
    '
    ' Parameter:
    ' string    The recurring token.
    ' -------------------------------------------------------------------------
    Public Function SetToken(token)
        m_path = "/checkout/recurring/" & token
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the uri of the resource.
    '
    ' Parameter/Returns:
    ' string    The uri.
    ' -------------------------------------------------------------------------
    Public Function GetLocation()
        GetLocation = m_location
    End Function

    Public Function SetLocation(uri)
        m_location = uri
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the content type of the resource.
    '
    ' Parameter/Returns:
    ' string    The content type.
    ' -------------------------------------------------------------------------
    Public Function GetContentType()
        GetContentType = m_contentType
    End Function

    Public Function SetContentType(ctype)
        m_contentType = ctype
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the accept media type of the resource.
    '
    ' Parameter/Returns:
    ' string    The accept media type.
    ' -------------------------------------------------------------------------
    Public Function GetAccept()
        GetAccept = m_accept
    End Function

    Public Function SetAccept(atype)
        m_accept = atype
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

    ' -------------------------------------------------------------------------
    ' Checks if an error occurred.
    '
    ' Returns:
    ' boolean    If an error occurred.
    ' -------------------------------------------------------------------------
    Public Function HasError()
        If (m_error Is Nothing) Then
            HasError = False
            Exit Function
        End If

        HasError = True
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the latest error.
    '
    ' Parameter/Returns:
    ' object    The ApiError object.
    ' -------------------------------------------------------------------------
    Public Function GetError()
        Set GetError = m_error
    End Function

    Public Function SetError(err)
        Set m_error = err
    End Function

    ' -------------------------------------------------------------------------
    ' Fetches the recurring status.
    ' -------------------------------------------------------------------------
    Public Function Fetch()
        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")
        options.Add "path", m_path

        Call m_connector.Apply("GET", Me, options)
    End Function

End Class

%>
