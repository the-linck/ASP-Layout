<%
BeforeView()

%><!DOCTYPE html>
<html>
<head><%
    if ViewData("charset") <> vbNullString then
        %><meta charset="<%= ViewData("charset") %>"><%
    end if
    %><meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">

    <title><%= ViewData("title") %></title><%

    GetScripts(ViewData("libraries").Items)
%></head>
<body><%
    Include(Join( _
        Array("Views/", ViewData("page"), ".view.asp") _
        , vbNullString _
    ))

    GetScripts(ViewData("plugins").Items)
%></body>
</html><%

AfterView()
%>