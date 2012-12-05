﻿<%
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

'------------------------------------------------------------------------------
' The HTTP request.
'------------------------------------------------------------------------------

Class Request
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_uri
    Private m_timeout

    ' -------------------------------------------------------------------------
    ' Gets or sets the uri.
    ' -------------------------------------------------------------------------
    Public Function GetUri()
        GetUri = m_uri
    End Function

    Public Function SetUri(uri)
        m_uri = uri
    End Function

    ' -------------------------------------------------------------------------
    ' Gets or sets the number of milliseconds before the connection times out.
    ' -------------------------------------------------------------------------
    Public Function GetTimeout()
        GetTimeout = m_timeout
    End Function

    Public Function SetTimeout(timeout)
        m_timeout = timeout
    End Function

End Class

%>