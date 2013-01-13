<%@ LANGUAGE="VBScript" CODEPAGE=65001 %>
<%
Option Explicit
Response.CodePage = 65001
Response.Charset = "UTF-8"
%>
<!-- #include virtual="/ASPUnit/include/ASPUnitRunner.asp" -->
<!-- #include file="../MockConnector.asp" -->
<!-- #include file="../MockHttpTransport.asp" -->
<!-- #include file="../../Klarna.Asp/JSON.asp" -->
<!-- #include file="../../Klarna.Asp/Order.asp" -->
<!-- #include file="../../Klarna.Asp/Digest.asp" -->
<!-- #include file="../../Klarna.Asp/UserAgent.asp" -->
<!-- #include file="../../Klarna.Asp/BasicConnector.asp" -->
<!-- #include file="../../Klarna.Asp/HttpRequest.asp" -->
<!-- #include file="../../Klarna.Asp/HttpResponse.asp" -->
<!-- #include file="../../Klarna.Asp/HttpTransport.asp" -->
<!-- #include file="DigestTest.asp" -->
<!-- #include file="OrderTest.asp" -->
<!-- #include file="OrderWithConnectorTest.asp" -->
<!-- #include file="UserAgentTest.asp" -->
<!-- #include file="BasicConnectorTest.asp" -->
<!-- #include file="BasicConnectorGetTest.asp" -->
<!-- #include file="BasicConnectorPostTest.asp" -->
<!-- #include file="HttpRequestTest.asp" -->
<!-- #include file="HttpResponseTest.asp" -->
<!-- #include file="HttpTransportTest.asp" -->
<%
' test runner
Dim runner
Set runner = New UnitRunner
Call runner.AddTestContainer(New DigestTest)
Call runner.AddTestContainer(New OrderTest)
Call runner.AddTestContainer(New OrderWithConnectorTest)
Call runner.AddTestContainer(New UserAgentTest)
Call runner.AddTestContainer(New BasicConnectorTest)
Call runner.AddTestContainer(New BasicConnectorGetTest)
Call runner.AddTestContainer(New BasicConnectorPostTest)
Call runner.AddTestContainer(New HttpRequestTest)
Call runner.AddTestContainer(New HttpResponseTest)
Call runner.AddTestContainer(New HttpTransportTest)

Call runner.Display()
%>
