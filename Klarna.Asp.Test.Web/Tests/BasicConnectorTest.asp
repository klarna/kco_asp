<%
'------------------------------------------------------------------------------
' Tests the BasicConnector class.
'------------------------------------------------------------------------------
Class BasicConnectorTest
    Public Function TestCaseNames()
        TestCaseNames = Array("UserAgent", "ApplyUrlInResource", "ApplyUrlInOptions")
    End Function

    Public Sub SetUp()
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the UserAgent property is correct.
    '--------------------------------------------------------------------------
    Public Sub UserAgent(testResult)
        Dim connector
        Set connector = new BasicConnector
        Dim ua
        Set ua = connector.GetUserAgent

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic", ua.ToString, "")

        ua.AddField "JS Lib", "jQuery", "1.8.2", Null

        Dim ua2
        Set ua2 = connector.GetUserAgent

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic JS Lib/jQuery_1.8.2", ua2.ToString, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses url in resource.
    '--------------------------------------------------------------------------
    Public Sub ApplyUrlInResource(testResult)
        Dim transport
        Set transport = new MockHttpTransport
        Dim digest
        Set digest = New Digest
        Dim connector
        Set connector = CreateBasicConnector(transport, digest, "My Secret")

        Dim order
        Set order = New Order
        order.SetLocation "http://klarna.com"

        Set transport.m_request = New HttpRequest

        Call connector.Apply("GET", order, Null)

        Call testResult.AssertEquals("http://klarna.com", transport.m_request.GetUri, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses url in options.
    '--------------------------------------------------------------------------
    Public Sub ApplyUrlInOptions(testResult)
        Dim transport
        Set transport = new MockHttpTransport
        Dim digest
        Set digest = New Digest
        Dim connector
        Set connector = CreateBasicConnector(transport, digest, "My Secret")

        Set transport.m_request = New HttpRequest

        Dim order
        Set order = New Order

        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")
        options.Add "url", "http://klarna.com"
        Call connector.Apply("GET", order, options)

        Call testResult.AssertEquals("http://klarna.com", transport.m_request.GetUri, "")
    End Sub
End Class

%>
