<%
'------------------------------------------------------------------------------
' Tests the HttpTransport class.
'------------------------------------------------------------------------------
Class HttpTransportTest
    Private transport

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation")
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

End Class

%>
