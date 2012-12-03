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
        Dim connector
        Set connector = new BasicConnector
        Dim ua
        Set ua = connector.GetUserAgent

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic", ua.ToString, "")

        ua.AddField "JS Lib", "jQuery", "1.8.2", Null

        Dim ua2
        Set ua2 = connector.GetUserAgent

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic JS Lib/jQuery_1.8.2", ua2.ToString, "")

    End Sub

End Class

%>
