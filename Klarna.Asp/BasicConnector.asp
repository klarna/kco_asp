<!-- #include file="json2.asp" -->

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

'--------------------------------------------------------------------------
' Creates a new instance or the BasicConnector class.
'
' Parameters:
' object    httpTransport    The HTTP transport to use.
' object    digest           The digest to use.
' string    secret           The secret to use.
'
' Returns:
' The basic connector instance.
'--------------------------------------------------------------------------
Public Function CreateBasicConnector(httpTransport, digest, secret)
    Dim connector
    Set connector = new BasicConnector
    connector.SetHttpTransport httpTransport
    connector.SetDigest digest
    connector.SetSecret secret

    Set CreateBasicConnector = connector
End Function

'------------------------------------------------------------------------------
' The basic connector.
'------------------------------------------------------------------------------
Class BasicConnector
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_userAgent
    Private m_httpTransport
    Private m_digest
    Private m_secret

    ' -------------------------------------------------------------------------
    ' Class constructor
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_userAgent = new UserAgent
    End Sub

    Private Sub Class_Terminate
    End Sub

    Public Sub SetHttpTransport(httpTransport)
        Set m_httpTransport = httpTransport
    End Sub

    Public Sub SetDigest(digest)
        Set m_digest = digest
    End Sub

    Public Sub SetSecret(secret)
        m_secret = secret
    End Sub

    ' -------------------------------------------------------------------------
    ' Gets or sets the user agent used for User-Agent header.
    '
    ' Parameter/Returns:
    ' object    The user agent.
    ' -------------------------------------------------------------------------
    Public Function GetUserAgent()
        Set GetUserAgent = m_userAgent
    End Function

    Public Function SetUserAgent(ua)
        Set m_userAgent = ua
    End Function

    ' -------------------------------------------------------------------------
    ' Applies a HTTP method on a specific resource.
    '
    ' Parameters:
    ' string    httpMethod      The HTTP method, GET or POST.
    ' object    order           The resource.
    ' object    options         The options.
    ' -------------------------------------------------------------------------
    Public Function Apply(httpMethod, order, options)
        Dim visitedUrl
        Set visitedUrl = Server.CreateObject("Scripting.Dictionary")

        Handle httpMethod, order, options, visitedUrl
    End Function

    ' -------------------------------------------------------------------------
    ' Handles a HTTP request.
    '
    ' Parameters:
    ' string    httpMethod      The HTTP method, GET or POST.
    ' object    order           The resource.
    ' object    options         The options.
    ' object    visitedUrl      List of visited url.
    '
    ' Returns:
    ' object    The response.
    ' -------------------------------------------------------------------------
    Private Function Handle(httpMethod, order, options, visitedUrl)
        Dim url
        url = GetUrl(order, options)

        Dim payLoad
        payLoad = ""
        If httpMethod = "POST" Then
            payLoad = GetData(order, options)
        End If

        Dim request
        Set request = CreateRequest(order, httpMethod, payLoad, url)

        Dim response
        Set response = m_httpTransport.Send(request)

        Set Handle = HandleResponse(response, httpMethod, order, visitedUrl)
    End Function

    ' -------------------------------------------------------------------------
    ' Gets the url to use, from options or resource.
    '
    ' Parameters:
    ' object    order           The resource.
    ' object    options         The options.
    '
    ' Returns:
    ' string    The url.
    ' -------------------------------------------------------------------------
    Private Function GetUrl(order, options)
        const URL_KEY = "url"

        Dim urlInOptions
        urlInOptions = False
        If IsObject(options) Then
            If options.Exists(URL_KEY) Then
                urlInOptions = True
            End If
        End If

        Dim url
        If urlInOptions Then
            url = options.Item(URL_KEY)
        Else
            url = order.GetLocation
        End If

        GetUrl = url
    End Function

    ' -------------------------------------------------------------------------
    ' Gets data to use, from options or resource.
    '
    ' Parameters:
    ' object    order           The resource.
    ' object    options         The options.
    '
    ' Returns:
    ' string    The data  in JSON format.
    ' -------------------------------------------------------------------------
    Private Function GetData(order, options)
        const DATA_KEY = "data"

        Dim dataInOptions
        dataInOptions = False
        If IsObject(options) Then
            If options.Exists(DATA_KEY) Then
                dataInOptions = True
            End If
        End If

        Dim data
        If dataInOptions Then
            Dim dictionaryData
            Set dictionaryData = options.Item(DATA_KEY)

            data = JSONEncodeDict("", dictionaryData)
        Else
            Dim marshalData
            Set marshalData = order.Marshal
            data = JSON.stringify(marshalData)
        End If

        GetData = data
    End Function

    ' -------------------------------------------------------------------------
    ' Creates a request.
    '
    ' Parameters:
    ' object    order           The resource.
    ' string    httpMethod      The HTTP method, GET or POST.
    ' string    payLoad         The payload.
    ' string    url             The url to use.
    '
    ' Returns:
    ' object    The request.
    ' -------------------------------------------------------------------------
    Private Function CreateRequest(order, httpMethod, payLoad, url)
        ' Create the request with correct method to use
        Dim request
        Set request = m_httpTransport.CreateRequest(url)
        request.SetMethod httpMethod

        ' Set HTTP Headers
        request.SetHeader "User-Agent", m_userAgent.ToString

        Dim digestString
        digestString = m_digest.Create(payLoad & m_secret)

        Dim authorization
        authorization = "Klarna " & digestString
        request.SetHeader "Authorization", authorization

        request.SetHeader "Accept", order.GetContentType

        If Len(payLoad) > 0 Then
            request.SetHeader "Content-Type", order.GetContentType
            request.SetData payLoad
        End If

        Set CreateRequest = request
    End Function

    ' -------------------------------------------------------------------------
    ' Handle response based on status.
    '
    ' Parameters:
    ' object    response        The response to handle.
    ' string    httpMethod      The HTTP method, GET or POST.
    ' object    order           The resource.
    ' object    visitedUrl      List of visited url.
    '
    ' Returns:
    ' object    The response.
    ' -------------------------------------------------------------------------
    Private Function HandleResponse(response, httpMethod, order, visitedUrl)
        VerifyResponse response

        Dim uri
        uri = response.GetHeader("Location")

        Dim statusCode
        statusCode = response.GetStatus()
        If statusCode = 200 Then        ' Update Data on resource.
            On Error resume Next

            Dim data
            data = response.GetData()
            If Len(data) > 0 Then
                order.Parse data
            End If

            If Err.Number <> 0 Then
                Err.Raise Err.Number, "Connector error", "Bad format on response content."
            End If
        ElseIf statusCode = 201 Then    ' Update location.
            order.SetLocation uri
        ElseIf statusCode = 301 Then    ' Update location and redirect if method is GET.
            order.SetLocation uri
            If httpMethod = "GET" Then
                Set HandleResponse = MakeRedirect(order, visitedUrl, uri)
                Exit Function
            End If
        ElseIf statusCode = 302 Then    ' Redirect if method is GET.
            If httpMethod = "GET" Then
                Set HandleResponse = MakeRedirect(order, visitedUrl, uri)
                Exit Function
            End If
        ElseIf statusCode = 303 Then    ' Redirect with GET, even if request is POST.
            Set HandleResponse = MakeRedirect(order, visitedUrl, uri)
            Exit Function
        End If

        Set HandleResponse = response
    End Function

    ' -------------------------------------------------------------------------
    ' Method to verify the response.
    '
    ' Parameters:
    ' object    response        The response to verify.
    ' -------------------------------------------------------------------------
    Private Sub VerifyResponse(response)
        Dim statusCode
        statusCode = response.GetStatus
        If statusCode >= 400 And statusCode <= 599 Then
            Err.Raise statusCode, "BasicConnector:VerifyResponse", "HTTP error code"
        End If
    End Sub

    ' -------------------------------------------------------------------------
    ' Makes a redirect.
    '
    ' Parameters:
    ' object    order           The resource.
    ' object    visitedUrl      List of visited url.
    ' string    url             The url to use.
    '
    ' Returns:
    ' object    The response.
    ' -------------------------------------------------------------------------
    Private Function MakeRedirect(order, visitedUrl, url)
        If visitedUrl.Exists(url) Then
            Err.Raise 5, "BasicConnector:MakeRedirect", "Infinite redirect loop detected."
        End If

        visitedUrl.Add url, ""

        Dim options
        Set options = Server.CreateObject("Scripting.Dictionary")
        options.Add "url", url

        Set MakeRedirect = Handle("GET", order, options, visitedUrl)
    End Function

End Class

%>