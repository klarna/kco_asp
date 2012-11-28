<!-- #include file="../../Klarna.Asp/jsonencode.asp" -->

<%
'------------------------------------------------------------------------------
' Tests the Order class.
'------------------------------------------------------------------------------
Class OrderTest
    Private m_connector
    Private m_order
    Private m_url
    Private m_contentType
    Private m_theInt
    Private m_theString
    Private m_theDate

    Public Function TestCaseNames()
        TestCaseNames = Array("ConstructionWithConnector", _
            "ContentType", "LocationEmpty", _
            "LocationSetGet", "ParseAndMarshal")
    End Function

    Public Sub SetUp()
        Set m_connector = new MockConnector
        Set m_order = CreateOrder(m_connector)
        m_url = "http://klarna.com"
        m_contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"
        m_theInt = 89
        m_theString = "A string"
        m_theDate = "2012-10-14"
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests the construction with connector.
    '--------------------------------------------------------------------------
    Public Sub ConstructionWithConnector(testResult)
        Call testResult.AssertEquals("", m_order.GetBaseUri, "BaseUri is empty")
        Call testResult.AssertEquals("", m_order.GetLocation, "Location is empty")
        Call testResult.AssertEquals("", m_order.GetContentType, "ContentType is empty")
        
        Dim data
        Set data = m_order.Marshal()
        Call testResult.AssertEquals(True, data Is Nothing, "Data is empty")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the content type is correct.
    '--------------------------------------------------------------------------
    Public Sub ContentType(testResult)
        Call testResult.AssertEquals("", m_order.GetContentType, "")
        m_order.SetContentType(m_contentType)
        Call testResult.AssertEquals(m_contentType, m_order.GetContentType, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that the location not is initialized.
    '--------------------------------------------------------------------------
    Public Sub LocationEmpty(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests set/get location.
    '--------------------------------------------------------------------------
    Public Sub LocationSetGet(testResult)
        Call testResult.AssertEquals("", m_order.GetLocation, "")
        m_order.SetLocation(m_url)
        Call testResult.AssertEquals(m_url, m_order.GetLocation, "")
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that parse and marshal works correctly.
    '--------------------------------------------------------------------------
    Public Sub ParseAndMarshal(testResult)
        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "TheInt", m_theInt
        data.Add "TheString", m_theString
        data.Add "TheDate", m_theDate

        Dim json
        json = JSONEncodeDict("", data)

        m_order.Parse(json)

        Dim resourceData
        Set resourceData = m_order.Marshal()

        Call testResult.AssertEquals(True, IsObject(resourceData), "")
        Call testResult.AssertEquals(m_theInt, resourceData.TheInt, "")
        Call testResult.AssertEquals(m_theString, resourceData.TheString, "")
        Call testResult.AssertEquals(m_theDate, resourceData.TheDate, "")

    End Sub

End Class

%>
