<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<%
'------------------------------------------------------------------------------
'   Copyright 2015 Klarna AB
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'
'   Klarna Support: support@klarna.com
'   http://developers.klarna.com/
'------------------------------------------------------------------------------
%>
<title>FetchRecurring.asp</title>
<!-- #include file="../Klarna.Asp/ApiError.asp" -->
<!-- #include file="../Klarna.Asp/JSON.asp" -->
<!-- #include file="../Klarna.Asp/RecurringStatus.asp" -->
<!-- #include file="../Klarna.Asp/Digest.asp" -->
<!-- #include file="../Klarna.Asp/UserAgent.asp" -->
<!-- #include file="../Klarna.Asp/BasicConnector.asp" -->
<!-- #include file="../Klarna.Asp/HttpRequest.asp" -->
<!-- #include file="../Klarna.Asp/HttpResponse.asp" -->
<!-- #include file="../Klarna.Asp/HttpTransport.asp" -->
<%
'------------------------------------------------------------------------------
' The fetch recurring status example.
'------------------------------------------------------------------------------
Class FetchRecurring

    '--------------------------------------------------------------------------
    ' The example.
    '--------------------------------------------------------------------------
    Public Sub Example()
        On Error Resume Next

        Dim sharedSecret : sharedSecret = "sharedSecret"
        Dim recurringToken : recurringToken = "ABC123"

        Dim connector : Set connector = CreateConnector(sharedSecret)
        connector.SetBaseUri KCO_TEST_BASE_URI

        Dim recurringStatus
        Set recurringStatus = CreateRecurringStatus(connector, recurringToken)

        recurringStatus.Fetch

        If recurringStatus.HasError = True Then
            Response.Write("Message: " & recurringStatus.GetError().Marshal().internal_message & "<br/>")
        End If

        If Err.Number <> 0 Then
            Response.Write("An error occurred: " & Err.Description)
            Err.Clear

            ' Error occurred, stop execution
            Exit Sub
        End If

        Dim resourceData
        Set resourceData = recurringStatus.Marshal()

        Response.Write("Payment method: " & resourceData.payment_method.type)

    End Sub

End Class

Dim example
Set example = New FetchRecurring
Call example.Example()

%>
