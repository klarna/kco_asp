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
    ' -------------------------------------------------------------------------
    Public Function GetUserAgent()
        Set GetUserAgent = m_userAgent
    End Function

    Public Function SetUserAgent(ua)
        Set m_userAgent = ua
    End Function

    ' -------------------------------------------------------------------------
    ' Applies a HTTP method on a specific resource.
    ' -------------------------------------------------------------------------
    Public Function Apply(httpMethod, order, options)
        Handle httpMethod, order, options, "new List<Uri>()"
    End Function

    ' -------------------------------------------------------------------------
    ' Handles a HTTP request.
    ' -------------------------------------------------------------------------
    Private Function Handle(httpMethod, order, options, visitedUrl)
        Dim url
        url = GetUrl(order, options)

        Dim payLoad
        payLoad = ""
        If httpMethod = "POST" Then
            payLoad = GetData(order, options)
        End If

        Dim hr
        Set hr = CreateRequest(order, httpMethod, payLoad, url)

        'Dim response
        'response = m_httpTransport.Send(request)

        'return HandleResponse(response, method, resource, visitedUrl);
    End Function

    ' -------------------------------------------------------------------------
    ' Gets the url to use, from options or resource.
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
    ' Data is in JSON format.
    ' -------------------------------------------------------------------------
    Private Function GetData(order, options)
        const DATA_KEY = "url"

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
            Set dictionaryData = options.Item(URL_KEY)

            data = JSONEncodeDict("", dictionaryData)
        Else
            data = order.Marshal
        End If

        GetData = data
    End Function

    ' -------------------------------------------------------------------------
    ' Creates a request.
    ' -------------------------------------------------------------------------
    Private Function CreateRequest(order, httpMethod, payLoad, url)
        ' Create the request with correct method to use
        Dim request
        Set request = m_httpTransport.CreateRequest(url)
        request.SetMethod httpMethod

        ' Set HTTP Headers
        'request.UserAgent = UserAgent.ToString();

        'var digestString = digest.Create(string.Concat(payLoad, secret));
        'var authorization = string.Format("Klarna {0}", digestString);
        'request.Headers.Add("Authorization", authorization);

        '    request.Accept = resource.ContentType;

        'if (payLoad.Length > 0)
        '{
        '    request.ContentType = resource.ContentType;
        '}

        Set CreateRequest = request
    End Function

End Class

%>