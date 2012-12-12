<%
'------------------------------------------------------------------------------
' Tests the HttpResponse class.
'------------------------------------------------------------------------------
Class HttpResponseTest
    Private hr

    Public Function TestCaseNames()
        TestCaseNames = Array("DefaultCreation", "Creation")
    End Function

    Public Sub SetUp()
        Set hr = New HttpResponse
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests default creation.
    '--------------------------------------------------------------------------
    Public Sub DefaultCreation(testResult)
        Call testResult.AssertEquals(0, hr.GetStatus, "The status")
        Dim h
        Set h = hr.GetHeaders()
        Call testResult.AssertEquals(True, IsObject(h), "The headers object")
        Call testResult.AssertEquals("", hr.GetData, "The data")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests creation.
    '--------------------------------------------------------------------------
    Public Sub Creation(testResult)
        Dim status
        status = 200

        Dim headers
        Set headers = Server.CreateObject("Scripting.Dictionary")
        headers.Add "Content-Type", "application/json"
        headers.Add "Accept-Charset", "utf-8"

        Dim data
        data = "{""Brand"":""Volvo""}"

        hr.Create status, headers, data

        Call testResult.AssertEquals(status, hr.GetStatus, "The status")

        Dim h
        Set h = hr.GetHeaders()
        Call testResult.AssertEquals(True, IsObject(h), "The headers object")
        Call testResult.AssertEquals(2, h.Count, "Number of headers")
        Call testResult.AssertEquals("application/json", hr.GetHeader("Content-Type"), "Header 1")
        Call testResult.AssertEquals("utf-8", hr.GetHeader("Accept-Charset"), "Header 2")

        Call testResult.AssertEquals(data, hr.GetData, "The data")
    End Sub

End Class

%>
