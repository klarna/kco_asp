<%
'------------------------------------------------------------------------------
' Tests the ApiError class
'------------------------------------------------------------------------------
Class ApiErrorTest
    Private m_error
    Private m_response

    Public Function TestCaseNames()
        TestCaseNames = Array("ParseMarshal", "GetSetResponse", "InitialParse", "InitialParseError")
    End Function

    Public Sub SetUp()
        Set m_response = New HttpResponse

        Set m_error = CreateApiError(m_response)
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the getter and setter for Response
    '--------------------------------------------------------------------------
    Public Sub GetSetResponse(testResult)
        Call testResult.AssertEquals(0, m_error.GetResponse().GetStatus(), "")

        Dim resp
        Set resp = new HttpResponse

        resp.Create 200, "", ""

        m_error.SetResponse resp

        Call testResult.AssertEquals(200, m_error.GetResponse().GetStatus(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Test the JSON parsing and serializing
    '--------------------------------------------------------------------------
    Public Sub ParseMarshal(testResult)
        Dim json, resourceData
        json = "{""key"":""val""}"

        Call m_error.Parse(json)

        Set resourceData = m_error.Marshal()

        Call testResult.AssertEquals("val", resourceData.key, "")
        Call testResult.AssertEquals(json, m_error.MarshalAsJson(), "")
    End Sub

    '--------------------------------------------------------------------------
    ' Test the JSON parsing when creating the ApiError object
    '--------------------------------------------------------------------------
    Public Sub InitialParse(testResult)
        Dim resp
        Set resp = new HttpResponse

        Dim status
        status = 200

        Dim headers
        headers = "Content-Type:application/json" & vbCrLf & _
                  "Accept-Charset:utf-8"  & vbCrLf & _
                  "Server: Microsoft-IIS/8.0" & vbCrLf & ":"

        Dim data
        data = "{""Brand"":""Volvo""}"

        resp.Create status, headers, data

        Dim error
        Set error = CreateApiError(resp)

        Call testResult.AssertEquals("Volvo", error.Marshal().Brand, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Test that an invalid JSON payload does not crash the code execution
    '--------------------------------------------------------------------------
    Public Sub InitialParseError(testResult)
        Dim resp
        Set resp = new HttpResponse

        Dim status
        status = 200

        Dim headers
        headers = "Content-Type:application/json" & vbCrLf & _
                  "Accept-Charset:utf-8"  & vbCrLf & _
                  "Server: Microsoft-IIS/8.0" & vbCrLf & ":"

        Dim data
        data = "{""Brand"",""Volvo""}"

        resp.Create status, headers, data

        Dim error
        Set error = CreateApiError(resp)

        Call testResult.Assert((error.Marshal() Is Nothing), "")
    End Sub

End Class

%>
