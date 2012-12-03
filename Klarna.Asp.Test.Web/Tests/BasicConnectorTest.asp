<%
'------------------------------------------------------------------------------
' Tests the BasicConnector class.
'------------------------------------------------------------------------------
Class BasicConnectorTest
    Public Function TestCaseNames()
        TestCaseNames = Array("UserAgent")
    End Function

    Public Sub SetUp()
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the UserAgent property is correct.
    '--------------------------------------------------------------------------
    Public Sub UserAgent(testResult)
        Call testResult.AssertEquals("", "", "")
    End Sub

End Class

%>
