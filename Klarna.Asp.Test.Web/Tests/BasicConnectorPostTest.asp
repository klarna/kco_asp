<%
'------------------------------------------------------------------------------
' Tests the BasicConnector class.
'------------------------------------------------------------------------------
Class BasicConnectorPostTest
    Private m_transport
    Private m_digest
    Private m_secret
    Private m_connector
    Private m_url
    Private m_contentType
    Private m_responseData

    Public Function TestCaseNames()
        TestCaseNames = Array("ApplyPost200", "ApplyPost200InvalidJson", _
            "ApplyPost201UpdatedLocation", "ApplyPost301EnsureNotRedirected", _
            "ApplyPost302EnsureNotRedirected", "ApplyPost303And200", _
            "ApplyPost303And503")
    End Function

    Public Sub SetUp()
        Set m_transport = new MockHttpTransport
        Set m_digest = New Digest
        m_secret = "My Secret"
        Set m_connector = CreateBasicConnector(m_transport, m_digest, m_secret)
        m_url = "http://klarna.com"
        m_contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"
        m_responseData = "{""Year"":2012}"
    End Sub

    Public Sub TearDown()
        Set m_transport = Nothing
        Set m_connector = Nothing
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 200 return.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost200(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse "{}"

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 200, "", m_responseData

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("POST", m_transport.m_requestInSend.GetMethod(), "")
        
        Call testResult.AssertEquals(m_connector.GetUserAgent().ToString(), _
            m_transport.m_requestInSend.GetHeader("User-Agent"), "")
        
        Dim digestString
        digestString = m_digest.Create("{}" & m_secret)
        Dim authorization
        authorization = "Klarna " & digestString
        Call testResult.AssertEquals(authorization, _
            m_transport.m_requestInSend.GetHeader("Authorization"), "")
        
        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Accept"), "")

        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Content-Type"), "")

        Dim orderData
        Set orderData = order.Marshal
        Call testResult.AssertEquals(m_responseData, JSON.stringify(orderData), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 200 return.
    ' But that invalid JSON in response raises an error.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost200InvalidJson(testResult)
        On Error Resume Next

        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse "{}"

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 200, "", "{{{{"

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("Bad format on response content.", Err.Description, "")

        Err.Clear()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 201 return.
    ' Verifies that location is updated.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost201UpdatedLocation(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse "{}"

        Set m_transport.m_request = New HttpRequest
        Dim updatedLocation
        updatedLocation = "http://NewLocation.com"
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 201, "Location:" & updatedLocation, m_responseData

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("POST", m_transport.m_requestInSend.GetMethod(), "")
        
        Call testResult.AssertEquals(m_connector.GetUserAgent().ToString(), _
            m_transport.m_requestInSend.GetHeader("User-Agent"), "")
        
        Dim digestString
        digestString = m_digest.Create("{}" & m_secret)
        Dim authorization
        authorization = "Klarna " & digestString
        Call testResult.AssertEquals(authorization, _
            m_transport.m_requestInSend.GetHeader("Authorization"), "")
        
        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Accept"), "")

        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Content-Type"), "")

        Call testResult.AssertEquals(updatedLocation, order.GetLocation, "")

        Dim orderData
        Set orderData = order.Marshal
        Call testResult.AssertEquals("{}", JSON.stringify(orderData), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 301 return.
    ' Verifies that no redirect is performed and that location is updated.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost301EnsureNotRedirected(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse m_responseData

        Dim newLocation
        newLocation = "http://NewLocation.com"

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 301, "Location:" & newLocation, ""
        Set m_transport.m_response2 = New HttpResponse
        m_transport.m_response2.Create 503, "", m_responseData

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("POST", m_transport.m_requestInSend.GetMethod(), "")
        
        Call testResult.AssertEquals(m_connector.GetUserAgent().ToString(), _
            m_transport.m_requestInSend.GetHeader("User-Agent"), "")
        
        Dim digestString
        digestString = m_digest.Create(m_responseData & m_secret)
        Dim authorization
        authorization = "Klarna " & digestString
        Call testResult.AssertEquals(authorization, _
            m_transport.m_requestInSend.GetHeader("Authorization"), "")
        
        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Accept"), "")

        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Content-Type"), "")

        Call testResult.AssertEquals(newLocation, order.GetLocation, "")

        Call testResult.AssertEquals(1, m_transport.m_responseCount, "")
    End Sub
        
    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 302 return.
    ' Verifies that no redirect is performed and that location NOT is updated.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost302EnsureNotRedirected(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse m_responseData

        Dim newLocation
        newLocation = "http://NewLocation.com"

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 302, "Location:" & newLocation, ""
        Set m_transport.m_response2 = New HttpResponse
        m_transport.m_response2.Create 503, "", m_responseData

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("POST", m_transport.m_requestInSend.GetMethod(), "")
        
        Call testResult.AssertEquals(m_connector.GetUserAgent().ToString(), _
            m_transport.m_requestInSend.GetHeader("User-Agent"), "")
        
        Dim digestString
        digestString = m_digest.Create(m_responseData & m_secret)
        Dim authorization
        authorization = "Klarna " & digestString
        Call testResult.AssertEquals(authorization, _
            m_transport.m_requestInSend.GetHeader("Authorization"), "")
        
        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Accept"), "")

        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Content-Type"), "")

        Call testResult.AssertEquals(m_url, order.GetLocation, "")

        Call testResult.AssertEquals(1, m_transport.m_responseCount, "")
    End Sub
        
    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 303 return and redirect to status 200.
    ' Verifies redirect, that redirect is with GET and that location NOT is updated.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost303And200(testResult)
        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse m_responseData

        Dim newLocation
        newLocation = "http://NewLocation.com"

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 303, "Location:" & newLocation, ""
        Set m_transport.m_response2 = New HttpResponse
        m_transport.m_response2.Create 200, "Location:" & m_url, m_responseData

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("GET", m_transport.m_requestInSend.GetMethod(), "")
        
        Call testResult.AssertEquals(m_connector.GetUserAgent().ToString(), _
            m_transport.m_requestInSend.GetHeader("User-Agent"), "")
        
        Dim digestString
        digestString = m_digest.Create("" & m_secret)
        Dim authorization
        authorization = "Klarna " & digestString
        Call testResult.AssertEquals(authorization, _
            m_transport.m_requestInSend.GetHeader("Authorization"), "")
        
        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Accept"), "")

        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Content-Type"), "")

        Dim orderData
        Set orderData = order.Marshal
        Call testResult.AssertEquals(m_responseData, JSON.stringify(orderData), "")

        Call testResult.AssertEquals(m_url, order.GetLocation, "")

        Call testResult.AssertEquals(2, m_transport.m_responseCount, "")
    End Sub
        
    '--------------------------------------------------------------------------
    ' Tests Apply with POST method and status 303 return and redirect to status 503.
    ' Verifies redirect, that redirect is with GET, that location NOT is updated
    ' and that an error is raised.
    '--------------------------------------------------------------------------
    Public Sub ApplyPost303And503(testResult)
        On Error Resume Next

        Dim order
        Set order = New Order
        order.SetLocation m_url
        order.SetContentType m_contentType
        order.Parse m_responseData

        Dim newLocation
        newLocation = "http://NewLocation.com"

        Set m_transport.m_request = New HttpRequest
        Set m_transport.m_response = New HttpResponse
        m_transport.m_response.Create 303, "Location:" & newLocation, ""
        Set m_transport.m_response2 = New HttpResponse
        m_transport.m_response2.Create 503, "Location:" & m_url, m_responseData

        Call m_connector.Apply("POST", order, Null)

        Call testResult.AssertEquals("GET", m_transport.m_requestInSend.GetMethod(), "")
        
        Call testResult.AssertEquals(m_connector.GetUserAgent().ToString(), _
            m_transport.m_requestInSend.GetHeader("User-Agent"), "")
        
        Dim digestString
        digestString = m_digest.Create("" & m_secret)
        Dim authorization
        authorization = "Klarna " & digestString
        Call testResult.AssertEquals(authorization, _
            m_transport.m_requestInSend.GetHeader("Authorization"), "")
        
        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Accept"), "")

        Call testResult.AssertEquals(m_contentType, _
            m_transport.m_requestInSend.GetHeader("Content-Type"), "")

        Dim orderData
        Set orderData = order.Marshal
        Call testResult.AssertEquals(m_responseData, JSON.stringify(orderData), "")

        Call testResult.AssertEquals(m_url, order.GetLocation, "")

        Call testResult.AssertEquals(2, m_transport.m_responseCount, "")

        Call testResult.AssertEquals(503, Err.Number, "")

        Err.Clear()
    End Sub
End Class

%>
