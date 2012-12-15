<%
'------------------------------------------------------------------------------
' Tests the HttpTransport class.
'------------------------------------------------------------------------------
Class HttpTransportTest
    Private transport

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation", "Timeout", "CreateRequest", _
            "SendReturningErrorCode")
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
        transport.SetTimeout 10000
        Call testResult.AssertEquals(10000, transport.GetTimeout, "")
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

    '--------------------------------------------------------------------------
    ' Tests Send returning error code.
    '--------------------------------------------------------------------------
    Public Sub SendReturningErrorCode(testResult)
        Dim errorCodes
        errorCodes = Array(400, 401, 402, 403, 404, 406, 409, 412, 415, 422, _
            428, 429, 500, 502, 503)

        Dim errorCode
        For Each errorCode in errorCodes
            Dim hr
            Set hr = transport.CreateRequest("http://httpbin.org/status/" & errorCode)
            hr.SetHeader "Content-Type", "application/xml"
            hr.SetHeader "Accept-Charset", "utf-8"

            Dim result
            Set result = transport.Send(hr)

            Call testResult.AssertEquals(errorCode, result.GetStatus, "")
            Call testResult.AssertEquals("", result.GetData, "")
        Next
    End Sub

End Class

%>
