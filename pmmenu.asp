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
<title>Preventive Maintenance Menu</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<style>
button
{
font-size:10pt;
width:140px;
height:25px
}
</style>
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Thursday, January 12, 2012
'	Creation.
'*************

dim HitCounts
Dim currentuser
Dim access
Dim access2

'Set/get hit counts.
HitCounts = HitCounter("pmmenu")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "pmmenu", currentuser)
If access <> "none" Then

%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td style="width:25%">&nbsp;</td>
				<td style="text-align:center;width:50%"><h1>Preventive Maintenance Menu</h1></td>
				<td style="text-align:right;vertical-align:top;width:25%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
<%
		access2 = UserAccess("inspections","pm_inspection",currentuser)
		If access2 <> "none" Then
%>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="pm_inspection" title="Enter/Edit Preventive Maintenance records for the current quarter" onclick="window.location='pm_inspection.asp'">Enter/Edit PMs</button></td>
				<td>&nbsp;</td>
			</tr>
<%
		End If
%>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="pm_manual_inspection" title="Print a blank PM data form" onclick="window.open('pm_manual_inspection.asp','PM');">Print PM Form</button></td>
				<td>&nbsp;</td>
			</tr>
<%
		access2 = UserAccess("inspections","pm_report_filter",currentuser)
		If access2 <> "none" Then
%>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="pm_report" title="Open the PM report" onclick="window.location='pm_report_filter.asp'">PM Report</button></td>
				<td>&nbsp;</td>
			</tr>
<%
		End If
%>
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
