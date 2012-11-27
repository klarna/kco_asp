<!-- #include file="json2.asp" -->

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
' Creates a new instance or the Order class.
'
' Parameters:
' object    connector    The connector to use.
'
' Returns:
' The order instance.
'--------------------------------------------------------------------------
Public Function CreateOrder(connector)
    Dim order
    Set order = new Order
    order.SetConnector(connector)

    Set CreateOrder = order
End Function

'------------------------------------------------------------------------------
' The order resource.
'------------------------------------------------------------------------------
Class Order
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_resourceData
    Private m_connector
    Private m_baseUri
    Private m_location
    Private m_contentType

    ' -------------------------------------------------------------------------
    ' Class constructor
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_resourceData = Nothing
        Set m_connector = Nothing
        m_baseUri = ""
        m_location = ""
        m_contentType = ""
    End Sub

    ' -------------------------------------------------------------------------
    ' Sets connector to use.
    ' -------------------------------------------------------------------------
    Public Sub SetConnector(connector)
        Set m_connector = connector
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the base uri that is used to create order resources.
    ' -------------------------------------------------------------------------
    Public Function GetBaseUri()
        GetBaseUri = m_baseUri
    End Function

    Public Function SetBaseUri(uri)
        m_baseUri = uri
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the uri of the resource.
    ' -------------------------------------------------------------------------
    Public Function GetLocation()
        GetLocation = m_location
    End Function

    Public Function SetLocation(uri)
        m_location = uri
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the content type of the resource.
    ' -------------------------------------------------------------------------
    Public Function GetContentType()
        GetContentType = m_contentType
    End Function

    Public Function SetContentType(uri)
        m_contentType = uri
    End Function

    ' -------------------------------------------------------------------------
    ' Replace resource with the new data.
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

End Class

%>
