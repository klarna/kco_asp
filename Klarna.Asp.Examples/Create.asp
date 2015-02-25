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
'[[examples-create]]
%>
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
' The create checkout example.
'------------------------------------------------------------------------------
Class Create

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
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

        ' Create connector
        Dim transport
        Set transport = new HttpTransport
        Dim digest
        Set digest = New Digest
        Dim sharedSecret
        sharedSecret = "sharedSecret"
        Dim connector
        Set connector = CreateBasicConnector(transport, digest, sharedSecret)

        Dim contentType
        contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"

        Dim eid
        eid = "0"

        Dim merchant
        Set merchant = Server.CreateObject("Scripting.Dictionary")
        merchant.Add "id", eid
        merchant.Add "terms_uri", "http://example.com/terms.asp"
        merchant.Add "checkout_uri", "https://example.com/checkout.asp"
        merchant.Add "confirmation_uri", _
            "https://example.com/thankyou.asp?sid=123&klarna_order={checkout.order.uri}"
        ' You cannot receive push notification on a non publicly available uri.
        merchant.Add "push_uri", _
            "https://example.com/push.asp?sid=123&klarna_order={checkout.order.uri}"

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "purchase_country", "SE"
        data.Add "purchase_currency", "SEK"
        data.Add "locale", "sv-se"
        data.Add "merchant", merchant
        data.Add "cart", cart

        Dim order
        Set order = CreateOrder(connector)
        order.SetBaseUri "https://checkout.testdrive.klarna.com/checkout/orders"
        order.SetContentType contentType

        order.Create data
    End Sub

End Class
'[[examples-create]]
 %>
