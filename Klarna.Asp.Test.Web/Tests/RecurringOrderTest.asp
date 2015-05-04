<%
'------------------------------------------------------------------------------
' Tests the RecurringOrder class
'------------------------------------------------------------------------------
Class RecurringOrderTest
    Private m_connector
    Private m_order
    Private m_token
    Private m_contentType
    Private m_accept

    Public Function TestCaseNames()
        TestCaseNames = Array("GetSetLocation", "GetSetContentType", "GetSetAccept", "GetSetError", "ParseMarshal", "Create")
    End Function

    Public Sub SetUp()
        m_token = "1234"
        m_contentType = "application/vnd.klarna.checkout.recurring-order-v1+json"
        m_accept = "application/vnd.klarna.checkout.recurring-order-accepted-v1+json"

        Set m_connector = new MockConnector

        Set m_order = CreateRecurringOrder(m_connector, m_token)
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for Location
    '--------------------------------------------------------------------------
    Public Sub GetSetLocation(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation(), "")

        m_order.SetLocation "http://test"

        Call testResult.AssertEquals("http://test", m_order.GetLocation(), "")
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
        Call testResult.AssertEquals(m_accept, m_order.GetAccept(), "")

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
        Dim json, resourceData
        json = "{""key"":""val""}"

        Call m_order.Parse(json)

        Set resourceData = m_order.Marshal()

        Call testResult.AssertEquals("val", resourceData.key, "")
        Call testResult.AssertEquals(json, m_order.MarshalAsJson(), "")
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
        Call testResult.AssertEquals("/checkout/recurring/1234/orders", m_connector.m_options.Item("path"), "")
        Call testResult.AssertEquals(data.Item("foo"), m_connector.m_options.Item("data").Item("foo"), "")
    End Sub

End Class

%>
