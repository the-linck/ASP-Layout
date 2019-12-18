<%' Minimalistic layout engine

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

' Clears variables after the View is executed, preparing the script to end.
' If you override this function, make sure to do this cleasing too.
Sub AfterView()
    set GET_ = Nothing
    set POST_ = Nothing
    set SERVER_ = Nothing

    ViewBag.RemoveAll() : set ViewBag = Nothing
    ViewData.RemoveAll() : set ViewData = Nothing
End Sub

' Anything to be executed right before the View is executed.
' This default implementation does nothing.
Sub BeforeView()
End Sub

' Prints all given CSS/JS scripts in HTML tags.
' @param {string|string[]|Dictionary} Files
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

' Gets the file name of the File on provided path.
' @param {string} File
' @return {string}
Function ScriptName(ByRef File)
    Dim BarIndex : BarIndex = InStrRev(File, "/")

    ScriptName = Mid(File, BarIndex + 1, LEN(File) - BarIndex)
End Function



' Displays the configured View file on current Template.
' If no Template is set, show the default blank HTML template.
' If the configured View doesn't exist, shows only the template.
Sub View()
    if ViewData("template") = vbNullString then ' Default Empty template
        ' Using server-side include because otherwise the path for default template would be undertemined
        ' And because it's faster, of course
        %><!--#include file="HtmlTemplate.asp"--><%
    else
        Include(ViewData("template"))
    end if
End Sub



' Settings of the layout engine
Dim ViewData  : set ViewData = CreateObject("Scripting.Dictionary")
' Scripts to be loaded on <head>.
Set ViewData("libraries")    = CreateObject("Scripting.Dictionary")
' Scripts to be loaded on <body>, after page's content.
Set ViewData("plugins")      = CreateObject("Scripting.Dictionary")



' PHP-named variables to access useful collections 
' Shorthand for Querystring variables
Dim GET_    : Set GET_    = Request.Querystring
' Shorthand for Formdata variables
Dim POST_   : Set POST_   = Request.Form
' Shorthand for Server constants
Dim SERVER_ : Set SERVER_ = Request.ServerVariables



' Template configuration
Dim LAYOUT_TMP
ViewData("charset")  = vbNullString
ViewData("no-cache") = False
' Performance: Temporary variable to avoid dictionary search
LAYOUT_TMP           = ScriptName(SERVER_("SCRIPT_NAME"))
ViewData("page")     = Replace(LAYOUT_TMP, ".asp", vbNullString)
ViewData("path")     = SERVER_("SCRIPT_NAME")
ViewData("script")   = LAYOUT_TMP
ViewData("template") = vbNullString
ViewData("title")    = vbNullString




' Arbitrary user data
Dim ViewBag : set ViewBag = CreateObject("Scripting.Dictionary")



' Performance: avoiding search dictionary multiple times
LAYOUT_TMP = Now()
' Library load timestamp
ViewData("ready-timestamp") = Join(Array(_
    Year(LAYOUT_TMP), "-", Month(LAYOUT_TMP), "-", Day(LAYOUT_TMP), " "_
    Hour(LAYOUT_TMP), ":", Minute(LAYOUT_TMP), ":", Second(LAYOUT_TMP) _
), vbNullString)
ViewData("ready") = LAYOUT_TMP
' Discarding temporary variable
LAYOUT_TMP = Empty %>