<%
'------------------------------------------------------------------------------
' Tests the BasicConnector class.
'------------------------------------------------------------------------------
Class BasicConnectorTest
    Private m_transport
    Private m_connector

    Public Function TestCaseNames()
        TestCaseNames = Array("UserAgent", "ApplyUrlInResource", "ApplyUrlInOptions")
    End Function

    Public Sub SetUp()
        Set m_transport = new MockHttpTransport
        Dim digest
        Set digest = New Digest
        Set m_connector = CreateBasicConnector(m_transport, digest, "My Secret")
    End Sub

    Public Sub TearDown()
        Set m_transport = Nothing
        Set m_connector = Nothing
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the UserAgent property is correct.
    '--------------------------------------------------------------------------
    Public Sub UserAgent(testResult)
        Dim ua
        Set ua = m_connector.GetUserAgent

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic", ua.ToString, "")

        ua.AddField "JS Lib", "jQuery", "1.8.2", Null

        Dim ua2
        Set ua2 = m_connector.GetUserAgent

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic JS Lib/jQuery_1.8.2", ua2.ToString, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses url in resource.
    '--------------------------------------------------------------------------
    Public Sub ApplyUrlInResource(testResult)
        Dim order
        Set order = New Order
        order.SetLocation "http://klarna.com"

        Set m_transport.m_request = New HttpRequest

        Call m_connector.Apply("GET", order, Null)

        Call testResult.AssertEquals("http://klarna.com", m_transport.m_request.GetUri, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses url in options.
    '--------------------------------------------------------------------------
    Public Sub ApplyUrlInOptions(testResult)
        Set m_transport.m_request = New HttpRequest

        Dim order
        Set order = New Order

        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")
        options.Add "url", "http://klarna.com"
        
        Call m_connector.Apply("GET", order, options)

        Call testResult.AssertEquals("http://klarna.com", m_transport.m_request.GetUri, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses data in resource.
    '--------------------------------------------------------------------------
    Public Sub ApplyDataInResource(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses data in options.
    '--------------------------------------------------------------------------
    Public Sub ApplyDataInOptions(testResult)
    End Sub

End Class

%>
