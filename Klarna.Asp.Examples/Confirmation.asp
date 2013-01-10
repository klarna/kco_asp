<%
'------------------------------------------------------------------------------
'   Copyright 2012 Klarna AB
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'       http://www.apache.org/licenses/LICENSE-2.0
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
' 
'   Klarna Support: support@klarna.com
'   http://integration.klarna.com/
'------------------------------------------------------------------------------
%>
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

        ' Create connector
        Dim transport
        Set transport = new HttpTransport
        Dim digest
        Set digest = New Digest
        Dim sharedSecret
        sharedSecret = "sharedSecret"
        Dim connector
        Set connector = CreateBasicConnector(transport, digest, sharedSecret)

        ' Retrieve location from session object.
        ' Use following in ASP.
        Dim checkoutId
        checkoutId = Session("klarna_checkout")
        Dim contentType
        contentType = "application/vnd.klarna.checkout.aggregated-order-v2+json"
        Dim order
        Set order = CreateOrder(connector)
        order.SetLocation checkoutId
        order.SetContentType contentType

        order.Fetch

        Dim resourceData
        Set resourceData = order.Marshal()
        If resourceData.status = "checkout_incomplete" Then
            ' Report error

            ' Use following in ASP.
            Response.Write("Checkout not completed, redirect to checkout.asp") 
        End If

        ' Display thank you snippet
        Dim gui
        Set gui = resourceData.gui
        Dim snippet
        snippet = gui.snippet

        ' DESKTOP: Width of containing block shall be at least 750px
        ' MOBILE: Width of containing block shall be 100% of browser
        ' window (No padding or margin)
        ' Use following in ASP.
        Response.Write("<div>" & snippet & "</div>")

        ' Clear session object.
        Session("klarna_checkout") = ""

        Err.Clear()
    End Sub

End Class
%>