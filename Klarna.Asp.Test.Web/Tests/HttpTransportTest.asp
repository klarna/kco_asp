<%
'------------------------------------------------------------------------------
' Tests the HttpTransport class.
'------------------------------------------------------------------------------
Class HttpTransportTest
    Private transport

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation", "Timeout", "CreateRequest")
    End Function

    Public Sub SetUp()
        Set transport = New HttpTransport
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests default creation.
    '--------------------------------------------------------------------------
    Public Sub Creation(testResult)
        Call testResult.AssertEquals(5000, transport.GetTimeout, "The timeout")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Timeout setter and getter.
    '--------------------------------------------------------------------------
    Public Sub Timeout(testResult)
        Call testResult.AssertEquals(5000, transport.GetTimeout, "")
        transport.SetTimeout 30000
        Call testResult.AssertEquals(30000, transport.GetTimeout, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests CreateRequest.
    '--------------------------------------------------------------------------
    Public Sub CreateRequest(testResult)
        Dim hr
        Set hr = transport.CreateRequest("http://klarna.com")
        Call testResult.AssertEquals("http://klarna.com", hr.GetUri, "")
        Call testResult.AssertEquals("GET", hr.GetMethod, "")
        Dim h
        Set h = hr.GetHeaders()
        Call testResult.AssertEquals(True, IsObject(h), "")
        Call testResult.AssertEquals("", hr.GetData, "")
    End Sub

End Class

%>
