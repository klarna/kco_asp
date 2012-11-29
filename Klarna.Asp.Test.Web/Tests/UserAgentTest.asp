<%
'------------------------------------------------------------------------------
' Tests the UserAgent class.
'------------------------------------------------------------------------------
Class UserAgentTest
    Private ua

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation")
    End Function

    Public Sub SetUp()
        Set ua = New UserAgent
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests Default UA string.
    '--------------------------------------------------------------------------
    Public Sub Creation(testResult)
        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic", _
                                        ua.ToString, "The UA string")
    End Sub

End Class

%>
