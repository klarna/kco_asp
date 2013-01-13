<!-- #include file="../../Klarna.Asp/json2.asp" -->

<%
'------------------------------------------------------------------------------
' Tests the Order class.
'------------------------------------------------------------------------------
Class OrderTest
    Private m_connector
    Private m_order
    Private m_url
    Private m_contentType
    Private m_theInt
    Private m_theString
    Private m_theDate

    Public Function TestCaseNames()
        TestCaseNames = Array("ConstructionWithConnector", _
            "ContentType", "LocationEmpty", _
            "LocationSetGet", "ParseAndMarshal", "ParseAndMarshalRealStructures")
    End Function

    Public Sub SetUp()
        Set m_connector = new MockConnector
        Set m_order = CreateOrder(m_connector)
        m_url = "http://klarna.com"
        m_contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"
        m_theInt = 89
        m_theString = "A string"
        m_theDate = "2012-10-14"
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the construction with connector.
    '--------------------------------------------------------------------------
    Public Sub ConstructionWithConnector(testResult)
        Call testResult.AssertEquals("", m_order.GetBaseUri, "BaseUri is empty")
        Call testResult.AssertEquals("", m_order.GetLocation, "Location is empty")
        Call testResult.AssertEquals("", m_order.GetContentType, "ContentType is empty")
        
        Dim data
        Set data = m_order.Marshal()
        Call testResult.AssertEquals(True, data Is Nothing, "Data is empty")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the content type is correct.
    '--------------------------------------------------------------------------
    Public Sub ContentType(testResult)
        Call testResult.AssertEquals("", m_order.GetContentType, "")
        m_order.SetContentType(m_contentType)
        Call testResult.AssertEquals(m_contentType, m_order.GetContentType, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the location not is initialized.
    '--------------------------------------------------------------------------
    Public Sub LocationEmpty(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests set/get location.
    '--------------------------------------------------------------------------
    Public Sub LocationSetGet(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation, "")
        m_order.SetLocation(m_url)
        Call testResult.AssertEquals(m_url, m_order.GetLocation, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that parse and marshal works correctly.
    '--------------------------------------------------------------------------
    Public Sub ParseAndMarshal(testResult)
        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "TheInt", m_theInt
        data.Add "TheString", m_theString
        data.Add "TheDate", m_theDate

        Dim jx
        Set jx = new JSONX
        Dim json
        json = jx.toJSON(Empty, data, true)

        m_order.Parse(json)

        Dim resourceData
        Set resourceData = m_order.Marshal()

        Call testResult.AssertEquals(True, IsObject(resourceData), "")
        Call testResult.AssertEquals(m_theInt, resourceData.TheInt, "")
        Call testResult.AssertEquals(m_theString, resourceData.TheString, "")
        Call testResult.AssertEquals(m_theDate, resourceData.TheDate, "")

    End Sub

    '--------------------------------------------------------------------------
    ' Tests that parse and marshal works correctly for real structures.
    '--------------------------------------------------------------------------
    Public Sub ParseAndMarshalRealStructures(testResult)
        ' Cart
        Dim item1
        Set item1 = Server.CreateObject("Scripting.Dictionary")
        item1.Add "quantity", 1
        item1.Add "reference", "BANAN01"
        item1.Add "name", "Bananana"
        item1.Add "unit_price", 450
        item1.Add "discount_rate", 0
        item1.Add "tax_rate", 2500

        Dim item2
        Set item2 = Server.CreateObject("Scripting.Dictionary")
        item2.Add "quantity", 1
        item2.Add "type", "shipping_fee"
        item2.Add "reference", "SHIPPING"
        item2.Add "name", "Shipping Fee"
        item2.Add "unit_price", 450
        item2.Add "discount_rate", 0
        item2.Add "tax_rate", 2500

        Dim cartItems(1)
        Set cartItems(0) = item1
        Set cartItems(1) = item2

        Dim cart
        Set cart = Server.CreateObject("Scripting.Dictionary")
        cart.Add "items", cartItems

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")

        Dim eid
        eid = "2"

        Dim merchant
        Set merchant = Server.CreateObject("Scripting.Dictionary")
        merchant.Add "id", eid
        merchant.Add "terms_uri", "http://localhost/terms.html"
        merchant.Add "checkout_uri", "http://localhost/checkout.asp"
        merchant.Add "confirmation_uri", "http://localhost/confirmation.asp"
        ' You cannot recieve push notification on a non publicly available uri.
        merchant.Add "push_uri", "http://localhost/push.asp"

        data.Add "purchase_country", "SE"
        data.Add "purchase_currency", "SEK"
        data.Add "locale", "sv-se"
        data.Add "merchant", merchant
        data.Add "cart", cart

        Dim jx
        Set jx = new JSONX
        Dim json
        json = jx.toJSON(Empty, data, true)

        m_order.Parse(json)

        Dim resourceData
        Set resourceData = m_order.Marshal()

        Call testResult.AssertEquals(True, IsObject(resourceData), "")
        Call testResult.AssertEquals("SE", resourceData.purchase_country, "")
        Call testResult.AssertEquals("SEK", resourceData.purchase_currency, "")
        Call testResult.AssertEquals("sv-se", resourceData.locale, "")

        Call testResult.AssertEquals("2", resourceData.merchant.id, "")
        Call testResult.AssertEquals("http://localhost/terms.html", resourceData.merchant.terms_uri, "")
        Call testResult.AssertEquals("http://localhost/checkout.asp", resourceData.merchant.checkout_uri, "")
        Call testResult.AssertEquals("http://localhost/confirmation.asp", resourceData.merchant.confirmation_uri, "")
        Call testResult.AssertEquals("http://localhost/push.asp", resourceData.merchant.push_uri, "")

        Call testResult.AssertEquals(1, resourceData.cart.items.[0].quantity, "")
        Call testResult.AssertEquals("BANAN01", resourceData.cart.items.[0].reference, "")
        Call testResult.AssertEquals("Bananana", resourceData.cart.items.[0].name, "")
        Call testResult.AssertEquals(450, resourceData.cart.items.[0].unit_price, "")
        Call testResult.AssertEquals(0, resourceData.cart.items.[0].discount_rate, "")
        Call testResult.AssertEquals(2500, resourceData.cart.items.[0].tax_rate, "")

        Call testResult.AssertEquals(1, resourceData.cart.items.[1].quantity, "")
        Call testResult.AssertEquals("shipping_fee", resourceData.cart.items.[1].type, "")
        Call testResult.AssertEquals("SHIPPING", resourceData.cart.items.[1].reference, "")
        Call testResult.AssertEquals("Shipping Fee", resourceData.cart.items.[1].name, "")
        Call testResult.AssertEquals(450, resourceData.cart.items.[1].unit_price, "")
        Call testResult.AssertEquals(0, resourceData.cart.items.[1].discount_rate, "")
        Call testResult.AssertEquals(2500, resourceData.cart.items.[1].tax_rate, "")

    End Sub
End Class

%>
