<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<script language="javascript">
function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Inspections Database Reports</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Wednesday, March 9, 2011
'   Creation
'*************

dim HitCounts
Dim currentuser
Dim access

'Set/get hit counts.
HitCounts = HitCounter("inspection_reports")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "reportmenu", currentuser)
If access <> "none" Then

%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td style="width:25%">&nbsp;</td>
				<td style="text-align:center;width:50%"><h1>Inspections Database Reports</h1></td>
				<td style="text-align:right;vertical-align:top;width:25%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><input type="button" name="psvinspectionsdue" value="PSV Inspections Due" title="Report of PSV inspections due to be performed" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='psv_inspections_due.asp'" /></td>
				<td>&nbsp;</td>
			</tr>
		</table>
	</div>
<%
Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</body>
</html>
