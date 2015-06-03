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
<title>CreateRecurring.asp</title>
<!-- #include file="../Klarna.Asp/ApiError.asp" -->
<!-- #include file="../Klarna.Asp/JSON.asp" -->
<!-- #include file="../Klarna.Asp/RecurringOrder.asp" -->
<!-- #include file="../Klarna.Asp/Digest.asp" -->
<!-- #include file="../Klarna.Asp/UserAgent.asp" -->
<!-- #include file="../Klarna.Asp/BasicConnector.asp" -->
<!-- #include file="../Klarna.Asp/HttpRequest.asp" -->
<!-- #include file="../Klarna.Asp/HttpResponse.asp" -->
<!-- #include file="../Klarna.Asp/HttpTransport.asp" -->
<%
'------------------------------------------------------------------------------
' The create recurring order example.
'------------------------------------------------------------------------------
Class CreateRecurring

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
        On Error Resume Next

        Dim eid : eid = "0"
        Dim sharedSecret : sharedSecret = "sharedSecret"
        Dim recurringToken : recurringToken = "ABC123"

        Dim connector : Set connector = CreateConnector(sharedSecret)
        connector.SetBaseUri KCO_TEST_BASE_URI

        Dim recurringOrder
        Set recurringOrder = CreateRecurringOrder(connector, recurringToken)

        Dim merchant
        Set merchant = Server.CreateObject("Scripting.Dictionary")
        merchant.Add "id", eid

        ' For testing purposes you can state either 'accept' or 'reject' at
        ' the end of the email addresses to trigger different responses,
        Dim email
        email = "checkout-se@testdrive.klarna.accept"

        Dim address
        Set address = Server.CreateObject("Scripting.Dictionary")
        address.Add "postal_code", "12345"
        address.Add "email", email
        address.Add "country", "se"
        address.Add "city", "Ankeborg"
        address.Add "family_name", "Approved"
        address.Add "given_name", "Testperson-se"
        address.Add "street_address", "Stårgatan 1"
        address.Add "phone", "070 111 11 11"

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
        item2.Add "discount_rate", 2500
        item2.Add "tax_rate", 2500

        Dim cartItems(1)
        Set cartItems(0) = item1
        Set cartItems(1) = item2

        Dim cart
        Set cart = Server.CreateObject("Scripting.Dictionary")
        cart.Add "items", cartItems

        Dim uniqueId
        uniqueId = "Some unique id..."

        Dim reference
        Set reference = Server.CreateObject("Scripting.Dictionary")
        reference.Add "orderid1", uniqueId

        ' If the order should be activated automatically.
        ' Set to True if you instead want a invoice created
        ' otherwise you will get a reservation.
        Dim activate
        activate = False

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "purchase_country", "SE"
        data.Add "purchase_currency", "SEK"
        data.Add "locale", "sv-se"
        data.Add "merchant", merchant
        data.Add "merchant_reference", reference
        data.Add "billing_address", address
        data.Add "shipping_address", address
        data.Add "cart", cart
        data.Add "activate", activate

        recurringOrder.Create data

        If recurringOrder.HasError = True Then
            Dim errData
            Set errData = recurringOrder.GetError().Marshal()

            If recurringOrder.GetError().GetResponse().GetStatus() = 402 Then
                Response.Write("Message: " & errData.reason & "<br/>")
            Else
                Response.Write("Message: " & errData.internal_message & "<br/>")
            End If
        End If

        If Err.Number <> 0 Then
            Response.Write("An error occurred: " & Err.Description)
            Err.Clear

            ' Error occurred, stop execution
            Exit Sub
        End If

        Dim resourceData
        Set resourceData = recurringOrder.Marshal()

        If activate = True Then
            Response.Write("Invoice number: " & resourceData.invoice)
        Else
            Response.Write("Reservation number: " & resourceData.reservation)
        End If

    End Sub

End Class

Dim example
Set example = New CreateRecurring
Call example.Example()

%>
