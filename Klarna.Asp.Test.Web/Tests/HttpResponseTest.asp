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
    End Sub

End Class

%>
