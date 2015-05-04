<%
'------------------------------------------------------------------------------
' Tests the RecurringStatus class
'------------------------------------------------------------------------------
Class RecurringStatusTest
    Private m_connector
    Private m_status
    Private m_token
    Private m_contentType
    Private m_accept

    Public Function TestCaseNames()
        TestCaseNames = Array("GetSetLocation", "GetSetContentType", "GetSetAccept", "GetSetError", "ParseMarshal", "Fetch")
    End Function

    Public Sub SetUp()
        m_token = "1234"
        m_contentType = "application/vnd.klarna.checkout.recurring-status-v1+json"

        Set m_connector = New MockConnector

        Set m_status = CreateRecurringStatus(m_connector, m_token)
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for Location
    '--------------------------------------------------------------------------
    Public Sub GetSetLocation(testResult)
        Call testResult.AssertEquals("", m_status.GetLocation(), "")

        m_status.SetLocation "http://test"

        Call testResult.AssertEquals("http://test", m_status.GetLocation(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for ContentType
    '--------------------------------------------------------------------------
    Public Sub GetSetContentType(testResult)
        Call testResult.AssertEquals(m_contentType, m_status.GetContentType(), "")

        m_status.SetContentType "test"

        Call testResult.AssertEquals("test", m_status.GetContentType(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for Accept
    '--------------------------------------------------------------------------
    Public Sub GetSetAccept(testResult)
        Call testResult.AssertEquals(m_accept, m_status.GetAccept(), "")

        m_status.SetAccept "test"

        Call testResult.AssertEquals("test", m_status.GetAccept(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter, setter and has error methods
    '--------------------------------------------------------------------------
    Public Sub GetSetError(testResult)
        Call testResult.AssertEquals(False, m_status.HasError(), "")

        Dim error
        Set error = Server.CreateObject("Scripting.Dictionary")
        error.Add "err", True

        Call m_status.SetError(error)

        Call testResult.AssertEquals(True, m_status.HasError(), "")
        Call testResult.AssertEquals(True, m_status.GetError().Item("err"), "")
    End Sub


    '--------------------------------------------------------------------------
    ' Test the JSON parsing and serializing
    '--------------------------------------------------------------------------
    Public Sub ParseMarshal(testResult)
        Dim json, resourceData
        json = "{""key"":""val""}"

        Call m_status.Parse(json)

        Set resourceData = m_status.Marshal()

        Call testResult.AssertEquals("val", resourceData.key, "")
        Call testResult.AssertEquals(json, m_status.MarshalAsJson(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Fetch works correctly.
    '--------------------------------------------------------------------------
    Public Sub Fetch(testResult)
        Call m_status.Fetch()

        Call testResult.AssertEquals("GET", m_connector.m_httpMethod, "")
        Call testResult.AssertEquals("/checkout/recurring/1234", m_connector.m_options.Item("path"), "")
    End Sub

End Class

%>
