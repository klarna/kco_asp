<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
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
%>
<!-- #include file="../Klarna.Asp/ApiError.asp" -->
<!-- #include file="../Klarna.Asp/JSON.asp" -->
<!-- #include file="../Klarna.Asp/Order.asp" -->
<!-- #include file="../Klarna.Asp/Digest.asp" -->
<!-- #include file="../Klarna.Asp/UserAgent.asp" -->
<!-- #include file="../Klarna.Asp/BasicConnector.asp" -->
<!-- #include file="../Klarna.Asp/HttpRequest.asp" -->
<!-- #include file="../Klarna.Asp/HttpResponse.asp" -->
<!-- #include file="../Klarna.Asp/HttpTransport.asp" -->
<%
'------------------------------------------------------------------------------
' The push example.
'------------------------------------------------------------------------------
Class Push

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
        On Error Resume Next

        Dim sharedSecret
        sharedSecret = "sharedSecret"

        ' Create connector
        Dim connector
        Set connector = CreateConnector(sharedSecret)

        ' Retrieve location from query string.
        ' Use following in ASP.
        Dim checkoutId
        checkoutId = Request.QueryString("checkout_uri")
        Dim order
        Set order = CreateOrder(connector)
        order.SetLocation checkoutId

        order.Fetch

        Dim resourceData
        Set resourceData = order.Marshal()
        If resourceData.status = "checkout_complete" Then
            ' At this point make sure the order is created in your
            ' system and send a confirmation email to the customer.

            Dim uniqueId
            uniqueId = "Some unique id..."

            Dim reference
            Set reference = Server.CreateObject("Scripting.Dictionary")
            reference.Add "orderid1", uniqueId

            Dim data
            Set data = Server.CreateObject("Scripting.Dictionary")
            data.Add "status", "created"
            data.Add "merchant_reference", reference

            order.Update data
        End If

        If order.HasError = True Then
            Response.Write("Message: " & order.GetError().Marshal().internal_message & "<br/>")
        End If

        If Err.Number <> 0 Then
            Response.Write("An error occurred: " & Err.Description)
            Err.Clear

            ' Error occurred, stop execution
            Exit Sub
        End If

        Set resourceData = order.Marshal()
        Response.Write(resourceData.status)

    End Sub

End Class

Dim example
Set example = New Push
Call example.Example()

%>
