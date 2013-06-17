<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Inspections Database</title>
<!--#include file="..\Functions\HitCounter.asp"-->
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<style>
button
{
font-size:10pt;
width:160px;
height:25px
}
</style>
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Monday, January 17, 2011
'   Creation
' Keith Brooks - Thursday, January 12, 2012
'	Added button for PMs.
'*************

dim HitCounts
Dim currentuser
Dim access
Dim access2

'Set/get hit counts.
HitCounts = HitCounter("inspection_home")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "default", currentuser)
If access <> "none" Then

%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td style="width:25%">&nbsp;</td>
				<td style="text-align:center;width:50%"><h1>Inspections Database</h1></td>
				<td style="text-align:right;vertical-align:top;width:25%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="rmp" title="Display RMP page"  onclick="window.location='rmp.asp'">RMP</button></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="enterinspection" title="Select an equipment item to enter, edit or print an inspection"  onclick="window.location='SelectItem.asp?form_action=enterinspection'">Inspections</button></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="reportmenu" title="Open the Inspections Database Report menu"  onclick="window.location='reportmenu.asp'">Reports</button></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="pmmenu" title="Open the Preventive Maintenance menu" onclick="window.location='pmmenu.asp'">Preventive Maintenance</button></td>
				<td>&nbsp;</td>
			</tr>
<%
		access2 = UserAccess("inspections","adminmenu",currentuser)
		If access2 <> "none" Then
%>
			<tr>
				<td>&nbsp;</td>
				<td style="text-align:center"><button type="button" name="adminmenu" title="Open the Inspections Database Administration menu"  onclick="window.location='adminmenu.asp'">Administration</button></td>
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
