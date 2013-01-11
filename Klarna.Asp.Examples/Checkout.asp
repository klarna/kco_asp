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
'   http://integration.klarna.com/
'------------------------------------------------------------------------------
'[[examples-checkout]]
%>
<!-- #include file="../Klarna.Asp/Order.asp" -->
<!-- #include file="../Klarna.Asp/Digest.asp" -->
<!-- #include file="../Klarna.Asp/UserAgent.asp" -->
<!-- #include file="../Klarna.Asp/BasicConnector.asp" -->
<!-- #include file="../Klarna.Asp/HttpRequest.asp" -->
<!-- #include file="../Klarna.Asp/HttpResponse.asp" -->
<!-- #include file="../Klarna.Asp/HttpTransport.asp" -->
<%
'------------------------------------------------------------------------------
' The checkout example.
'------------------------------------------------------------------------------
Class Checkout

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
        On Error Resume Next

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

        Dim order
        Set order = Nothing

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")

        ' Retrieve location from session object.
        Dim resourceUri
        resourceUri = Session("klarna_checkout")
        If resourceUri <> "" Then
            Set order = CreateOrder(connector)
            order.SetLocation resourceUri
            order.SetContentType contentType

            order.Fetch

            ' Reset cart
            data.Add "cart", cart

            order.Update data

            If Err.Number <> 0 Then
                ' Reset session
                Set order = Nothing
                Session("klarna_checkout") = ""
            End If
        End If

        If order Is Nothing Then
            ' Start a new session

            Dim eid
            eid = "0"

            Dim merchant
            Set merchant = Server.CreateObject("Scripting.Dictionary")
            merchant.Add "id", eid
            merchant.Add "terms_uri", "http://localhost/terms.html"
            merchant.Add "checkout_uri", "http://localhost/checkout.asp"
            merchant.Add "confirmation_uri", "http://localhost/confirmation.asp"
            ' You cannot recieve push notification on a non publicly available uri.
            merchant.Add "push_uri", "http://localhost/push.asp"

            data.RemoveAll()
            data.Add "purchase_country", "SE"
            data.Add "purchase_currency", "SEK"
            data.Add "locale", "sv-se"
            data.Add "merchant", merchant
            data.Add "cart", cart

            Set order = CreateOrder(connector)
            order.SetBaseUri "https://klarnacheckout.apiary.io/checkout/orders"
            order.SetContentType contentType

            order.Create data
            order.Fetch
        End If

        ' Store location of checkout session is session object.
        Session("klarna_checkout") = order.GetLocation

        ' Display checkout
        Dim resourceData
        Set resourceData = order.Marshal()
        Dim gui
        Set gui = resourceData.gui
        Dim snippet
        snippet = gui.snippet

        ' DESKTOP: Width of containing block shall be at least 750px
        ' MOBILE: Width of containing block shall be 100% of browser
        ' window (No padding or margin)
        ' Use following in ASP.
        Response.Write("<div>" & snippet & "</div>")

        Err.Clear()
    End Sub

End Class
'[[examples-checkout]]
%>
