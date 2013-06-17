<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<script language="javascript">
function openhelp() {
 window.open("Equipment Inspection Database Administrators Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Inspections Database Administration</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Wednesday, February 16, 2011
'   Creation
'*************

dim HitCounts
Dim currentuser
Dim access
Dim access2

'Set/get hit counts.
HitCounts = HitCounter("inspections_admin")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
'access = UserAllowed(currentuser, "masterbatchentry")
access = UserAccess("inspections", "adminmenu", currentuser)
If access <> "none" Then

%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td style="text-align:left;vertical-align:top;width:33%"><a href="default.asp">Home</a></td>
				<td style="text-align:center;width:34%">&nbsp;</td>
				<td style="text-align:right;vertical-align:top;width:33%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
			<tr>
				<td colspan="3" style="text-align:center"><h1 />Inspections Database Administration</td>
			</tr>
			<tr>
				<td style="text-align:center;vertical-align:top;border-style:solid;border-color:darkblue;border-width:2px">
					<table style="width:100%">
						<tr>
							<td style="text-align:center;font-weight:bold;font-size:14pt">Tools</td>
						</tr>
<%
		access2 = UserAccess("inspections","audittrail",currentuser)
		If access2 <> "none" Then
%>
						<tr>
							<td style="text-align:center"><input type="button" name="audittrail" value="Inspections&#10;Audit Trail" title="Open the inspections audit trail query form" style="font-size:10pt;width:140px;height:50px"  onclick="window.location='audittrail.asp'" /></td>
						</tr>
<%
		Else
%>
						<tr>
							<td>&nbsp;</td>
						</tr>
<%
		End If
%>
<%
		access2 = UserAccess("inspections","adminaudittrail",currentuser)
		If access2 <> "none" Then
%>
						<tr>
							<td style="text-align:center"><input type="button" name="adminaudittrail" value="Administration&#10;Audit Trail" title="Open the admin audit trail query form" style="font-size:10pt;width:140px;height:50px"  onclick="window.location='adminaudittrail.asp'" /></td>
						</tr>
<%
		Else
%>
						<tr>
							<td>&nbsp;</td>
						</tr>
<%
		End If
%>
					</table>
				</td>
				<td style="text-align:center;vertical-align:top;border-style:solid;border-color:darkblue;border-width:2px">
					<table style="width:80%">
						<tr>
							<td style="text-align:center;font-weight:bold;font-size:14pt">Lookup Tables</td>
						</tr>
						<tr>
<%
		access2 = UserAccess("inspections","adminmenu",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" id="equipmenttypes" name="equipmenttypes" value="Equipment Types" title="Maintain the equipment types" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='equipmenttypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("inspections","adminmenu",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" id="equipmentitems" name="equipmentitems" value="Equipment Items" title="Maintain the equipment items" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='equipmentitems.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("inspections","adminmenu",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" id="technicaldata" name="technicaldata" value="Technical Data" title="Maintain the technical data for an equipment item" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='SelectItem.asp?form_action=edittechnicaldata'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
					</table>
				</td>
				<td style="text-align:center;vertical-align:top;border-style:solid;border-color:darkblue;border-width:2px">
					<table style="width:100%">
						<tr>
							<td style="text-align:center;font-weight:bold;font-size:14pt">Security</td>
						</tr>
						<tr>
<%
		access2 = UserAccess("inspections","rolemembers",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="rolemembers" value="Role Members" title="Assign users to security roles" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='rolemembers.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("inspections","useraccess",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="useraccess" value="User Access" title="Assign user privileges for application forms" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='useraccess.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
					</table>
				</td>
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
