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
<title>Confirmation.asp</title>
<!-- #include file="../Klarna.Asp/ApiError.asp" -->
<!-- #include file="../Klarna.Asp/JSON.asp" -->
<!-- #include file="../Klarna.Asp/Order.asp" -->
<!-- #include file="../Klarna.Asp/Digest.asp" -->
<!-- #include file="../Klarna.Asp/UserAgent.asp" -->
<!-- #include file="../Klarna.Asp/BasicConnector.asp" -->
<!-- #include file="../Klarna.Asp/HttpRequest.asp" -->
<!-- #include file="../Klarna.Asp/HttpResponse.asp" -->
<!-- #include file="../Klarna.Asp/HttpTransport.asp" -->
<%
'------------------------------------------------------------------------------
' The confirmation example.
'------------------------------------------------------------------------------
Class Confirmation

    '--------------------------------------------------------------------------
    ' This example demonstrates the use of the Klarna library to complete
    ' the purchase and display the confirmation page snippet.
    '--------------------------------------------------------------------------
    Public Sub Example()
        On Error Resume Next

        Dim sharedSecret : sharedSecret = "sharedSecret"
        ' Retrieve location from session object.
        ' Use following in ASP.
        Dim orderID : orderID = Session("klarna_order_id")

        Dim connector: Set connector = CreateConnector(sharedSecret)
        connector.SetBaseUri KCO_TEST_BASE_URI

        Dim order : Set order = CreateOrder(connector)
        order.ID orderID

        order.Fetch

        If order.HasError = True Then
            Response.Write("Message: " & order.GetError().Marshal().internal_message & "<br/>")
        End If

        If Err.Number <> 0 Then
            Response.Write("An error occurred: " & Err.Description)
            Err.Clear

            ' Error occurred, stop execution
            Exit Sub
        End If

        Dim resourceData : Set resourceData = order.Marshal()
        If resourceData.status = "checkout_incomplete" Then
            ' Report error

            ' Use following in ASP.
            Response.Write("Checkout not completed, redirect to checkout.asp")

            Exit Sub
        End If

        ' Display thank you snippet
        Dim snippet : snippet = resourceData.gui.snippet

        ' DESKTOP: Width of containing block shall be at least 750px
        ' MOBILE: Width of containing block shall be 100% of browser
        ' window (No padding or margin)
        ' Use following in ASP.
        Response.Write("<div>" & snippet & "</div>")

        ' Clear session object.
        Session("klarna_checkout") = ""

    End Sub

End Class

Dim example
Set example = New Confirmation
Call example.Example()

%>
