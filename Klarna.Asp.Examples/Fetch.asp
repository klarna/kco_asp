<%
'------------------------------------------------------------------------------
' The fetch checkout example.
'------------------------------------------------------------------------------
Class Fetch

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
        ' Create connector
        Dim transport
        Set transport = new HttpTransport
        Dim digest
        Set digest = New Digest
        Dim sharedSecret
        sharedSecret = "sharedSecret"
        Dim connector
        Set connector = CreateBasicConnector(transport, digest, sharedSecret)

        Dim resourceUri
        resourceUri = "https://checkout.testdrive.klarna.com/checkout/orders/ABC123"

        Dim contentType
        contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"
        Dim order
        Set order = CreateOrder(connector)
        order.SetLocation resourceUri
        order.SetContentType contentType

        order.Fetch
    End Sub

End Class
%>