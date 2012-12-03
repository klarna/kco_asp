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

'------------------------------------------------------------------------------
' The user agent string creation class.
'------------------------------------------------------------------------------
Class UserAgent
    ' -------------------------------------------------------------------------
    ' Private members
    ' -------------------------------------------------------------------------
    Private m_fields

    ' -------------------------------------------------------------------------
    ' Class constructor
    '
    ' Initializes a new instance of the UserAgent class.
    ' Following fields are predefined:
    ' Library and Language.
    ' -------------------------------------------------------------------------
    Private Sub Class_Initialize
        Set m_fields = Server.CreateObject("Scripting.Dictionary")

        AddField "Library", "Klarna.ApiWrapper", "1.0", Null
        AddField "Language", "ASP", "Classic", Null
    End Sub

    Private Sub Class_Terminate
        Set m_fields = Nothing
    End Sub

    ' -------------------------------------------------------------------------
    ' Adds a field to the field collection.
    ' -------------------------------------------------------------------------
    Public Sub AddField(field, name, version, options)
        If m_fields.Exists(field) Then
            Err.Raise 457, "UserAgent:AddField", "Field already exists."
        End If

        Dim optionsString
        optionsString = ""

        If IsArray(options) Then
            optionsString = " " & Join(options, ";")
        End If

        Dim item
        item = field & "/" & name & "_" & version & optionsString

        m_fields.Add field, item
    End Sub

    ' -------------------------------------------------------------------------
    ' Returns the user agent string.
    ' -------------------------------------------------------------------------
    Public Function ToString()
        ToString = Join(m_fields.Items, " ")
    End Function
End Class
%>