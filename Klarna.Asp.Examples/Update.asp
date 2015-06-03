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
<title>Update.asp</title>
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
' The update checkout example.
'------------------------------------------------------------------------------
Class Update

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
        On Error Resume Next

        Dim sharedSecret : sharedSecret = "sharedSecret"
        Dim orderID : orderID = "ABC123"

        Dim connector : Set connector = CreateConnector(sharedSecret)
        connector.SetBaseUri KCO_TEST_BASE_URI

        Dim order : Set order = CreateOrder(connector)
        order.ID orderID

        ' Cart
        Dim item1
        Set item1 = Server.CreateObject("Scripting.Dictionary")
        item1.Add "reference", "123456789"
        item1.Add "name", "Klarna t-shirt"
        item1.Add "quantity", 2
        item1.Add "unit_price", 12300
        item1.Add "discount_rate", 1000
        item1.Add "tax_rate", 2500

        Dim item2
        Set item2 = Server.CreateObject("Scripting.Dictionary")
        item2.Add "type", "shipping_fee"
        item2.Add "reference", "SHIPPING"
        item2.Add "name", "Shipping Fee"
        item2.Add "quantity", 1
        item2.Add "unit_price", 4900
        item2.Add "tax_rate", 2500

        Dim cartItems(1)
        Set cartItems(0) = item1
        Set cartItems(1) = item2

        Dim cart
        Set cart = Server.CreateObject("Scripting.Dictionary")
        cart.Add "items", cartItems

        ' Reset cart
        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "cart", cart

        order.Update data

        If order.HasError = True Then
            Response.Write("Message: " & order.GetError().Marshal().internal_message & "<br/>")
        End If

        If Err.Number <> 0 Then
            Response.Write("An error occurred: " & Err.Description)
            Err.Clear

            ' Error occurred, stop execution
            Exit Sub
        End If

    End Sub

End Class

Dim example
Set example = New Update
Call example.Example()

%>
