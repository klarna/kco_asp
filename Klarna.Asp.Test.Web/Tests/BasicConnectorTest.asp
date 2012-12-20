<%
'------------------------------------------------------------------------------
' Tests the BasicConnector class.
'------------------------------------------------------------------------------
Class BasicConnectorTest
    Private m_transport
    Private m_connector
    Private m_url

    Public Function TestCaseNames()
        TestCaseNames = Array("UserAgent", "ApplyUrlInResource", "ApplyUrlInOptions", _
            "ApplyDataInResource", "ApplyDataInOptions")
    End Function

    Public Sub SetUp()
        Set m_transport = new MockHttpTransport
        Dim digest
        Set digest = New Digest
        Set m_connector = CreateBasicConnector(m_transport, digest, "My Secret")
        m_url = "http://klarna.com"
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
        order.SetLocation m_url

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 200, "", ""

        Call m_connector.Apply("GET", order, Null)

        Call testResult.AssertEquals("http://klarna.com", m_transport.m_request.GetUri, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses url in options.
    '--------------------------------------------------------------------------
    Public Sub ApplyUrlInOptions(testResult)
        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 200, "", ""

        Dim order
        Set order = New Order

        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")
        options.Add "url", m_url
        
        Call m_connector.Apply("GET", order, options)

        Call testResult.AssertEquals("http://klarna.com", m_transport.m_request.GetUri, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses data in resource.
    '--------------------------------------------------------------------------
    Public Sub ApplyDataInResource(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        Dim jsonData
        jsonData = "{""Year"":2012}"
        m_transport.m_response.Create 200, "", jsonData

        order.Parse jsonData
        
        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals(jsonData, m_transport.m_requestInSend.GetData, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that Apply uses data in options.
    '--------------------------------------------------------------------------
    Public Sub ApplyDataInOptions(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        Dim jsonData
        jsonData = "{""Year"":2012}"
        m_transport.m_response.Create 200, "", jsonData

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "Year", 2012
        
        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")
        options.Add "data", data

        Call m_connector.Apply("POST", order, options)

        Call testResult.AssertEquals(jsonData, m_transport.m_requestInSend.GetData, "")
    End Sub

End Class

%>
