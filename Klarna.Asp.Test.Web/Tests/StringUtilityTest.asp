<!-- #include virtual="/includes/StringUtility.inc.asp" -->
<%
Class StringUtilityTest
	Private m_stringUtility

	Public Function TestCaseNames()
		TestCaseNames = Array("ToLower_should_convert_uppercase_to_lowercase", "ToLower_should_leave_lowercase_the_same")
	End Function

	Public Sub SetUp()
		Set m_stringUtility = New StringUtility
	End Sub

	Public Sub TearDown()
	End Sub

	Public Sub ToLower_should_convert_uppercase_to_lowercase(testResult)
		Dim actual
		actual = m_stringUtility.ToLower("ABC")
		Call testResult.AssertEquals("abc", actual, "")
	End Sub

	Public Sub ToLower_should_leave_lowercase_the_same(testResult)
		Dim actual
		actual = m_stringUtility.ToLower("abc")
		Call testResult.AssertEquals("abc", actual, "")
	End Sub
End Class
%>
