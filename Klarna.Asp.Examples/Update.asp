<%
'------------------------------------------------------------------------------
' The update checkout example.
'------------------------------------------------------------------------------
Class Update

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

        Dim resourceUri
        resourceUri = "https://checkout.testdrive.klarna.com/checkout/orders/ABC123"
        Dim order
        Set order = CreateOrder(connector)
        order.SetLocation resourceUri
        order.SetContentType contentType

        ' Reset cart
        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "cart", cart

        order.Update data
    End Sub

End Class
 %>