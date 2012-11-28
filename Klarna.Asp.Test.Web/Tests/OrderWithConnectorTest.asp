<%
'------------------------------------------------------------------------------
' Tests the Order class with Connector.
'------------------------------------------------------------------------------
Class OrderWithConnectorTest
    Private m_url
    Private m_contentType
    Private m_connector
    Private m_order

    Public Function TestCaseNames()
        TestCaseNames = Array("Create", "CreateAlternativeEntryPoint", _
            "Fetch", "Update")
    End Function

    Public Sub SetUp()
        m_url = "http://klarna.com"
        m_contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"
        Set m_connector = new MockConnector
        Set m_order = CreateOrder(m_connector)
        m_order.SetBaseUri(m_url)
        m_order.SetContentType(m_contentType)
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Create works correctly.
    '--------------------------------------------------------------------------
    Public Sub Create(testResult)
        m_connector.m_httpMethod = ""
        Set m_connector.m_order = Nothing
        Set m_connector.m_options = Nothing

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "foo", "boo"

        Call m_order.Create(data)

        Call testResult.AssertEquals("POST", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals(m_url, m_connector.m_order.GetBaseUri, "")
        Call testResult.AssertEquals(m_order.GetBaseUri, m_connector.m_options.Item("url"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the Create with alternative entry point works correctly.
    '--------------------------------------------------------------------------
    Public Sub CreateAlternativeEntryPoint(testResult)
        m_connector.m_httpMethod = ""
        Set m_connector.m_order = Nothing
        Set m_connector.m_options = Nothing

        m_order.SetBaseUri("https://checkout.klarna.com/beta/checkout/orders")

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "foo", "boo"

        Call m_order.Create(data)

        Call testResult.AssertEquals("POST", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals(m_order.GetBaseUri, m_connector.m_order.GetBaseUri, "")
        Call testResult.AssertEquals(m_order.GetBaseUri, m_connector.m_options.Item("url"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Fetch works correctly.
    '--------------------------------------------------------------------------
    Public Sub Fetch(testResult)
        m_connector.m_httpMethod = ""
        Set m_connector.m_order = Nothing
        Set m_connector.m_options = Nothing

        m_order.SetLocation("http://klarna.com/foo/bar/15")

        Call m_order.Fetch()

        Call testResult.AssertEquals("GET", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals(m_order.GetLocation, m_connector.m_order.GetLocation, "")
        Call testResult.AssertEquals(m_order.GetLocation, m_connector.m_options.Item("url"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Update works correctly.
    '--------------------------------------------------------------------------
    Public Sub Update(testResult)
        m_connector.m_httpMethod = ""
        Set m_connector.m_order = Nothing
        Set m_connector.m_options = Nothing

        m_order.SetLocation("http://klarna.com/foo/bar/15")

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "foo", "boo"

        Call m_order.Update(data)

        Call testResult.AssertEquals("POST", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals(m_url, m_connector.m_order.GetBaseUri, "")
        Call testResult.AssertEquals(m_order.GetLocation, m_connector.m_options.Item("url"), "")
    End Sub
End Class
%>