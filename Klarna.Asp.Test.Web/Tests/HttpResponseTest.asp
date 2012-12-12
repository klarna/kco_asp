<%
'------------------------------------------------------------------------------
' Tests the HttpResponse class.
'------------------------------------------------------------------------------
Class HttpResponseTest
    Private hr

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation")
    End Function

    Public Sub SetUp()
        Set hr = New HttpResponse
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests default creation.
    '--------------------------------------------------------------------------
    Public Sub Creation(testResult)
        Call testResult.AssertEquals(0, hr.GetStatus, "The status")
        Dim h
        Set h = hr.GetHeaders()
        Call testResult.AssertEquals(True, IsObject(h), "The headers object")
        Call testResult.AssertEquals("", hr.GetData, "The data")
    End Sub

End Class

%>
