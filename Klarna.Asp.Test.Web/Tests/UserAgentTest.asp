<%
'------------------------------------------------------------------------------
' Tests the UserAgent class.
'------------------------------------------------------------------------------
Class UserAgentTest
    Private ua

    Public Function TestCaseNames()
        TestCaseNames = Array("Creation", "AddFieldWithoutOptions", _
            "AddFieldWithOptions", "CannotRedefineField")
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

    '--------------------------------------------------------------------------
    ' Test to add a field without options.
    '--------------------------------------------------------------------------
    Public Sub AddFieldWithoutOptions(testResult)
        ua.AddField "JS Lib", "jQuery", "1.8.2", Null

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic JS Lib/jQuery_1.8.2", _
                                        ua.ToString, "The UA string")
    End Sub

    '--------------------------------------------------------------------------
    ' Test to add a field with options.
    '--------------------------------------------------------------------------
    Public Sub AddFieldWithOptions(testResult)
        Dim options
        options = Array("LanguagePack/7", "JsLib/2.0")
        ua.AddField "Module", "Magento", "5.0,", options

        Call testResult.AssertEquals("Library/Klarna.ApiWrapper_1.0 Language/ASP_Classic Module/Magento_5.0, LanguagePack/7;JsLib/2.0", _
                                        ua.ToString, "The UA string")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that redefinition of field throws an exception.
    '--------------------------------------------------------------------------
    Public Sub CannotRedefineField(testResult)
        On Error Resume Next
        
        ua.AddField "Library", "None", "0.0,", Null
        
        Call testResult.AssertEquals(457, Err.Number, "Redefine Library")
        Err.Clear()

        ua.AddField "Language", "None", "0.0,", Null

        Call testResult.AssertEquals(457, Err.Number, "Redefine Library")
        Err.Clear()
    End Sub
End Class

%>
