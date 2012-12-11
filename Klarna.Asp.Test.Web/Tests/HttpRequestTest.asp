<%
'------------------------------------------------------------------------------
' Tests the HttpRequest class.
'------------------------------------------------------------------------------
Class HttpRequestTest
    Private hr

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation", "Uri", "Method", "Header", "Data")
    End Function

    Public Sub SetUp()
        Set hr = New HttpRequest
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests default creation.
    '--------------------------------------------------------------------------
    Public Sub Creation(testResult)
        Call testResult.AssertEquals("", hr.GetUri, "The uri")
        Call testResult.AssertEquals("GET", hr.GetMethod, "The method")
        Dim h
        Set h = hr.GetHeaders()
        Call testResult.AssertEquals(True, IsObject(h), "The headers object")
        Call testResult.AssertEquals("", hr.GetData, "The data")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Uri setter and getter.
    '--------------------------------------------------------------------------
    Public Sub Uri(testResult)
        Call testResult.AssertEquals("", hr.GetUri, "")
        hr.SetUri "http://klarna.com"
        Call testResult.AssertEquals("http://klarna.com", hr.GetUri, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Method setter and getter.
    '--------------------------------------------------------------------------
    Public Sub Method(testResult)
        Call testResult.AssertEquals("GET", hr.GetMethod, "")
        hr.SetMethod "POST"
        Call testResult.AssertEquals("POST", hr.GetMethod, "")
        hr.SetMethod "gEt"
        Call testResult.AssertEquals("GET", hr.GetMethod, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Header setter and getters.
    '--------------------------------------------------------------------------
    Public Sub Header(testResult)
        Call testResult.AssertEquals("", hr.GetHeader("Content-Type"), "")
        hr.SetHeader "Content-Type", "application/json"
        Call testResult.AssertEquals("application/json", hr.GetHeader("Content-Type"), "")
        hr.SetHeader "Content-Type", "application/xml"
        Call testResult.AssertEquals("application/xml", hr.GetHeader("Content-Type"), "")

        hr.SetHeader "Accept-Charset", "utf-8"
        Dim headers
        Set headers = hr.GetHeaders()
        Call testResult.AssertEquals(2, headers.Count, "")
        Call testResult.AssertEquals("application/xml", headers.Item("Content-Type"), "")
        Call testResult.AssertEquals("utf-8", headers.Item("Accept-Charset"), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Data setter and getter.
    '--------------------------------------------------------------------------
    Public Sub Data(testResult)
        Call testResult.AssertEquals("", hr.GetData, "")
        hr.SetData "{""Brand"":""Volvo""}"
        Call testResult.AssertEquals("{""Brand"":""Volvo""}", hr.GetData, "")
    End Sub

End Class

%>
