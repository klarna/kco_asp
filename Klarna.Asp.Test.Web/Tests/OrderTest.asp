<%
'------------------------------------------------------------------------------
' Tests the Order class.
'------------------------------------------------------------------------------
Class OrderTest
    Public Function TestCaseNames()
        TestCaseNames = Array("ConstructionWithConnector", _
            "ConstructionWithResourceUri", "ContentType", "LocationNull", _
            "LocationSetGet", "Parse", "Marshal", "ValuesGet")
    End Function

    Public Sub SetUp()
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the construction with connector.
    '--------------------------------------------------------------------------
    Public Sub ConstructionWithConnector(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the construction with resource uri.
    '--------------------------------------------------------------------------
    Public Sub ConstructionWithResourceUri(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the content type is correct.
    '--------------------------------------------------------------------------
    Public Sub ContentType(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the location not is initialized.
    '--------------------------------------------------------------------------
    Public Sub LocationNull(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests set/get location.
    '--------------------------------------------------------------------------
    Public Sub LocationSetGet(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that parse works correctly.
    '--------------------------------------------------------------------------
    Public Sub Parse(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that marshal works correctly.
    '--------------------------------------------------------------------------
    Public Sub Marshal(testResult)
    End Sub

    '--------------------------------------------------------------------------
    ' Tests get values.
    '--------------------------------------------------------------------------
    Public Sub ValuesGet(testResult)
    End Sub

End Class

%>
