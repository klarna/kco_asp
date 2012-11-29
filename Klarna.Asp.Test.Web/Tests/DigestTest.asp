<!-- #include file="../../Klarna.Asp/jsonencode.asp" -->

<%
'------------------------------------------------------------------------------
' Tests the Digest class.
'------------------------------------------------------------------------------
Class DigestTest
    Public Function TestCaseNames()
        TestCaseNames = Array("CreateDigest")
    End Function

    Public Sub SetUp()
    End Sub

    Public Sub TearDown()
    End Sub

    '--------------------------------------------------------------------------
    ' Tests that creation of digest string is correct.
    '--------------------------------------------------------------------------
    Public Sub CreateDigest(testResult)
        Dim article
        Set article = Server.CreateObject("Scripting.Dictionary")
        article.Add "artno", "id_1"
        article.Add "name", "product"
        article.Add "price", 12345
        article.Add "vat", 25
        article.Add "qty", 1

        Dim goodsList(0)
        set goodsList(0) = article

        Dim data
        Set data = Server.CreateObject("Scripting.Dictionary")
        data.Add "eid", 1245
        data.Add "goods_list", goodsList
        data.Add "currency", "SEK"
        data.Add "country", "SWE"
        data.Add "language", "SV"

        Dim json
        json = JSONEncodeDict("", data)

        Dim digest
        Set digest = New Digest
        Dim actual
        actual = digest.Create(json & "mySecret")
        
        Dim expected
        expected = "MO/6KvzsY2y+F+/SexH7Hyg16gFpsPDx5A2PtLZd0Zs="
        
        Call testResult.AssertEquals(expected, actual, "The digest string")
    End Sub

End Class

%>
