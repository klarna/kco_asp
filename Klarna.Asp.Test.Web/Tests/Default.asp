<%@ LANGUAGE="VBScript" CODEPAGE=65001 %>
<%
Option Explicit
Response.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/ASPUnit/include/ASPUnitRunner.asp" -->
<!-- #include file="DigestTest.asp" -->
<!-- #include file="OrderTest.asp" -->
<%
' test runner
Dim runner
Set runner = New UnitRunner
Call runner.AddTestContainer(New DigestTest)
Call runner.AddTestContainer(New OrderTest)

Call runner.Display()
%>
