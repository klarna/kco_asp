<!-- #include file="../../Klarna.Asp/Digest.asp" -->
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
        Dim digest
        Set digest = New Digest

        Dim json
        json = "message"
        
        Dim actual
        actual = digest.Create(json)
        
        Dim expected
        expected = "q1MKE+RZFJgrefm34/uplM/R8/si9xzqGvvwK0YMbR0="
        
        Call testResult.AssertEquals(expected, actual, "The digest string")
    End Sub

'------------------------------------------------------------------------------
' TODO: Sync to testcase in .Net
'------------------------------------------------------------------------------
'   var article = new Dictionary<string, object>()
'   {
'       { "artno", "id_1" }, 
'       { "name", "product" }, 
'       { "price", 12345 }, 
'       { "vat", 25 }, 
'       { "qty", 1 }
'   };

'   var goodsList = new List<Dictionary<string, object>>() { article };

'   var data = new Dictionary<string, object>()
'   {
'       { "eid", 1245 }, 
'       { "goods_list", goodsList }, 
'       { "currency", "SEK" }, 
'       { "country", "SWE" }, 
'       { "language", "SV" }
'   };

'   var json = JsonConvert.SerializeObject(data);

'   var actual = digest.Create(string.Concat(json, "mySecret"));
'------------------------------------------------------------------------------

End Class
%>
