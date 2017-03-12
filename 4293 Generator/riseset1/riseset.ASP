<%@ Language="VBScript" %>
<% Option Explicit %>
<html>

<head>
<title>Sample usage of RiseSet.DLL</title>
</head>
<%
' ############################################
' Riseset.DLL must be available and registered
' on the webserver before this will work.
' The call to the DLL is
' MyClass.SRS(latitude, longitude, timezone, RiseSet, dateyear, datemonth, dateday)
' where all variables are numbers, except RiseSet which is:
' R or S (for Sunrise or Sunset)
' MR or MS (for Moonrise or Moonset)
' NR or NS (for Nautical twilight rise/set)
' ############################################
%>

<body>
<%	Dim MyClass
	Set MyClass = CreateObject("RiseSetDLL.Class")
%>

<p><small><font face="Arial" color="#0000FF">Sunrise time: <%=MyClass.SRS(-35, 149, 11, "R", year(now), month(now), day(now))%> <br>
Sunset time : <%=MyClass.SRS(-35, 149, 11, "S", year(now), month(now), day(now))%> <br>
Moonrise time: <%=MyClass.SRS(-35, 149, 11, "MR", year(now), month(now), day(now))%> <br>
Moonset time : <%=MyClass.SRS(-35, 149, 11, "MS", year(now), month(now), day(now))%> <br>
Nautical twilight (rise) time: <%=MyClass.SRS(-35, 149, 11, "NR", year(now), month(now), day(now))%> <br>
Nautical twilight (set) time : <%=MyClass.SRS(-35, 149, 11, "NS", year(now), month(now), day(now))%> <br>
</font><small><font face="Arial">© Copyright 2002 Michael Pottinger (http://mhpnet.com)</font></small></small></p>
<%Set MyClass = Nothing%>
</body>
</html>
