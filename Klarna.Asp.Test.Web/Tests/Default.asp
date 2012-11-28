<%@ LANGUAGE="VBScript" CODEPAGE=65001 %>
<%
Option Explicit
Response.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/ASPUnit/include/ASPUnitRunner.asp" -->
<!-- #include file="../MockConnector.asp" -->
<!-- #include file="../../Klarna.Asp/Order.asp" -->
<!-- #include file="DigestTest.asp" -->
<!-- #include file="OrderTest.asp" -->
<!-- #include file="OrderWithConnectorTest.asp" -->
<%
' test runner
Dim runner
Set runner = New UnitRunner
Call runner.AddTestContainer(New DigestTest)
Call runner.AddTestContainer(New OrderTest)
Call runner.AddTestContainer(New OrderWithConnectorTest)

Call runner.Display()
%>
