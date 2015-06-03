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
        TestCaseNames = Array("ID", "GetSetLocation", "GetSetContentType", "GetSetAccept", _
            "GetSetError", "ParseMarshal","ParseAndMarshalRealStructures", "Create", _
            "Fetch", "Update")
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

    Public Sub ID(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation(), "")

        m_order.ID "1234"

        Call testResult.AssertEquals("http://stub.com/checkout/orders/1234", m_order.GetLocation(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for Location
    '--------------------------------------------------------------------------
    Public Sub GetSetLocation(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation(), "")

        m_order.SetLocation m_url

        Call testResult.AssertEquals(m_url, m_order.GetLocation(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for ContentType
    '--------------------------------------------------------------------------
    Public Sub GetSetContentType(testResult)
        Call testResult.AssertEquals(m_contentType, m_order.GetContentType(), "")

        m_order.SetContentType "test"

        Call testResult.AssertEquals("test", m_order.GetContentType(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for Accept
    '--------------------------------------------------------------------------
    Public Sub GetSetAccept(testResult)
        Call testResult.AssertEquals("", m_order.GetAccept(), "")

        m_order.SetAccept "test"

        Call testResult.AssertEquals("test", m_order.GetAccept(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter, setter and has error methods
    '--------------------------------------------------------------------------
    Public Sub GetSetError(testResult)
        Call testResult.AssertEquals(False, m_order.HasError(), "")

        Dim error
        Set error = Server.CreateObject("Scripting.Dictionary")
        error.Add "err", True

        Call m_order.SetError(error)

        Call testResult.AssertEquals(True, m_order.HasError(), "")
        Call testResult.AssertEquals(True, m_order.GetError().Item("err"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Test the JSON parsing and serializing
    '--------------------------------------------------------------------------
    Public Sub ParseMarshal(testResult)
        Call testResult.AssertEquals(True, m_order.Marshal() Is Nothing, "Data is empty")

        Dim json, resourceData
        json = "{""key"":""val""}"

        Call m_order.Parse(json)

        Set resourceData = m_order.Marshal()

        Call testResult.AssertEquals("val", resourceData.key, "")
        Call testResult.AssertEquals(json, m_order.MarshalAsJson(), "")
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

    '--------------------------------------------------------------------------
    ' Tests that Create works correctly.
    '--------------------------------------------------------------------------
    Public Sub Create(testResult)
        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "foo", "boo"

        Call m_order.Create(data)

        Call testResult.AssertEquals("POST", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals("/checkout/orders", m_connector.m_options.Item("path"), "")
        Call testResult.AssertEquals(data.Item("foo"), m_connector.m_options.Item("data").Item("foo"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Fetch works correctly.
    '--------------------------------------------------------------------------
    Public Sub Fetch(testResult)
        Call m_order.SetLocation(m_url)

        Call m_order.Fetch()

        Call testResult.AssertEquals("GET", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals(m_url, m_connector.m_options.Item("url"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Update works correctly.
    '--------------------------------------------------------------------------
    Public Sub Update(testResult)
        Call m_order.SetLocation(m_url)

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "foo", "boo"

        Call m_order.Update(data)

        Call testResult.AssertEquals("POST", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals(m_url, m_connector.m_options.Item("url"), "")
        Call testResult.AssertEquals(data.Item("foo"), m_connector.m_options.Item("data").Item("foo"), "")
    End Sub

End Class

%>
