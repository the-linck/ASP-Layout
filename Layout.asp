<!--#include file="DynamicInclude.asp"--><%
' Adds scripts to be loaded as a libraries (on <head>).
' @param {string|string[]} Scripts
Sub AddLibraries(ByRef Scripts)
    Dim Libraries : set Libraries = ViewData("libraries")
    Dim Index : Index = Libraries.Count

    Dim Script
    if IsArray(Scripts) then
        For Each Script in Scripts
            ' Checking type to prevent errors
            if VarType(Script) = vbString then
                Libraries(Index) = Script

                Index = Index + 1
            end if
        Next
    else
        Script = Scripts

        Libraries(Index) = Script
    end if
End Sub
' Adds scripts to be loaded as a plugins (on <body>, after page's content).
' @param {string|string[]} Scripts
Sub AddPlugins(ByRef Scripts)
    Dim Libraries : set Libraries = ViewData("plugins")
    Dim Index : Index = Libraries.Count

    Dim Script
    if IsArray(Scripts) then
        For Each Script in Scripts
            ' Checking type to prevent errors
            if VarType(Script) = vbString then
                Libraries(Index) = Script

                Index = Index + 1
            end if
        Next
    else
        Script = Scripts

        Libraries(Index) = Script
    end if
End Sub

' Clears variables after the View is executed, preparing the sript to end.
' If you overide this function, make sure to do this cleasing too.
Sub AfterView()
    set GET_ = Nothing
    set POST_ = Nothing
    set SERVER_ = Nothing

    ViewBag.RemoveAll() : set ViewBag = Nothing
    ViewData.RemoveAll() : set ViewData = Nothing
End Sub

' Functionality to be executed right before the View.
' This default implementation does nothing.
Sub BeforeView()
End Sub

Sub GetScripts(ByRef Files)
    Dim Extension
    Dim File
    Dim Iterate
    Dim Path

    if ViewData("no-cache") then
        Path = Array(vbNullString, "?TimeTag=", ViewData("ready-time"))
    else
        Path = Array(vbNullString)
    end if

    if IsArray(Files) then
        Iterate = Files
    else
        Select Case VarType(Files)
            Case vbObject
                if TypeName(Files) = "Dictionary" then
                    Iterate = Files.Items
                else
                    Err.Raise 13
                end if
            Case vbString
                Iterate = Array(Files)
            Case Else
                Err.Raise 13
        End Select
    end if

    For Each File in Iterate
        Extension = LCase( _
            Right(File, Len(File) - (InStrRev(File, ".") + 1))
        )
        Path(0) = File

        Select Case Extension
            Case "css"
                %><link href="<%= Join(Path, vbNullString) %>" rel="stylesheet" /><%
            Case "js"
                %><script src="<%= Join(Path, vbNullString) %>"></script><%
        End Select
    Next
End Sub

Function ScriptName(ByRef File)
    Dim BarIndex : BarIndex = InStrRev(File, "/")

    ScriptName = Mid(File, BarIndex + 1, LEN(File) - BarIndex)
End Function



Sub View()
    ' Using server-side include because otherwise the path for default template would be undertemined
    ' And because it's faster, of course
    %><!--#include file="DynamicInclude.asp"--><%
End Sub



' Configuration collection
Dim ViewData  : set ViewData = CreateObject("Scripting.Dictionary")
Set ViewData("libraries")    = CreateObject("Scripting.Dictionary")
Set ViewData("plugins")      = CreateObject("Scripting.Dictionary")

ViewData("no-cache")         = False

' PHP-named variables to access useful collections 
Dim GET_    : Set GET_    = Request.Querystring
Dim POST_   : Set POST_   = Request.Form
Dim SERVER_ : Set SERVER_ = Request.ServerVariables

' Template configuration
ViewData("charset")  = vbNullString
ViewData("template") = vbNullString
ViewData("path")     = SERVER_("SCRIPT_NAME")
ViewData("script")   = ScriptName(SERVER_("SCRIPT_NAME"))
ViewData("title")    = vbNullString

ViewData("page")    = Replace(ViewData("script"), ".asp", vbNullString)

' User data collection
Dim ViewBag : set ViewBag = CreateObject("Scripting.Dictionary")

' Library loaded timestamp
ViewData("ready-time")     = Now() %>