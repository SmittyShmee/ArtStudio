<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Gerald Steadman Smith Art Studio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="stylesheet" HREF="gsas_styles.css" TYPE="text/css">
</head>

<SCRIPT LANGUAGE=JAVASCRIPT>
<!-- Hide script from old browers
if (document.images) {

bullet_on = new Image
bullet_off = new Image
bullet_on.src = "images/bullet_on.gif"
bullet_off.src = "images/bullet_off.gif"

imgoff = new Image(); imgoff.src = "images/lg_start.gif"; 
img1_1on = new Image(); img1_1on.src = "images/lg_yves_larocque.jpg";
img1_2on = new Image(); img1_2on.src = "images/lg_martine_chartrand.jpg";
img2_1on = new Image(); img2_1on.src = "images/lg_the_red_wall.jpg";
img2_2on = new Image(); img2_2on.src = "images/lg_the_circle_game.jpg";
img2_3on = new Image(); img2_3on.src = "images/lg_a_touch_of_class.jpg";
img2_4on = new Image(); img2_4on.src = "images/lg_on_target.jpg";
img2_5on = new Image(); img2_5on.src = "images/lg_jack_and_jill.jpg";
img2_6on = new Image(); img2_6on.src = "images/lg_the_language_of_art.jpg";
img3_1on = new Image(); img3_1on.src = "images/lg_clarks_harbour_wharf.jpg";
img3_2on = new Image(); img3_2on.src = "images/lg_strong_westerly_gale.jpg";
img3_3on = new Image(); img3_3on.src = "images/lg_western_head_light.jpg";
img3_4on = new Image(); img3_4on.src = "images/lg_passing_seal_island_light.jpg";
img3_5on = new Image(); img3_5on.src = "images/lg_swimms_point_wharf.jpg";
img3_6on = new Image(); img3_6on.src = "images/lg_hawk_point_beach.jpg";
img4_1on = new Image(); img4_1on.src = "images/lg_floating_circle2.jpg";
img4_2on = new Image(); img4_2on.src = "images/lg_floating_circle4.jpg";

}
else {

}

// Highlight Image on Mouseover
function hiLite(imgName,imgObjName) {
	if (document.images) {
		document.images[imgName].src = eval(imgObjName + ".src");
	}
}


function sp(pic)
{
	window.open(pic,'picture','width=640,height=584,top=20,left=20,toolbars=no,resizable=1');
}

function validateEmail(email, msg, optional) 
{
if (!email.value && optional) {
	return true;
}
var re_mail = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+        ([a-zA-Z])+$/;
if (!re_mail.test(email.value)) {
	alert(msg);
	email.focus();
	email.select();
	return false;
}
return true;
}

// end hiding script from old browsers -->
</script>

<body>

<div align="center"><table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/template_heading_r1_c1.gif" width="640" height="114" alt=""></td>
  </tr>
  <tr>
    <td><img src="images/template_heading_r2_c1.gif" width="640" height="92	"></td>
  </tr>
  <tr>
    <td><table width="100%" bgcolor="#2E78B9" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="20" align="center"><img src="images/bullet_off.gif" name="menu_item_1" width="16" height="19" alt="" border="0"></td>
		<td width="119" align="center"><a href="about_the_artist.html" onMouseOver="document.menu_item_1.src=bullet_on.src;"
						onMouseOut="document.menu_item_1.src=bullet_off.src">ABOUT</a></td>
        <td width="47" align="center"><img src="images/bullet_off.gif" name="menu_item_2" width="16" height="19" alt="" border="0"></td>
        <td width="130" align="center"><a href="gallery.html" onMouseOver="document.menu_item_2.src=bullet_on.src;"
						onMouseOut="document.menu_item_2.src=bullet_off.src">GALLERY</a></td>
        <td width="52" align="center"><img src="images/bullet_off.gif" name="menu_item_3" width="16" height="19" alt="" border="0"></td>
        <td width="130" align="center"><a href="contact.asp" onMouseOver="document.menu_item_3.src=bullet_on.src;"
						onMouseOut="document.menu_item_3.src=bullet_off.src">CONTACT</a></td>
        <td width="20" align="center"><img src="images/bullet_off.gif" name="menu_item_4" width="16" height="19" alt="" border="0"></td>
        <td width="102" align="center"><a href="index.html" onMouseOver="document.menu_item_4.src=bullet_on.src;"
						onMouseOut="document.menu_item_4.src=bullet_off.src">HOME</a></td>
        <td width="20" align="center"><img src="images/bullet_off.gif" width="16" height="19" alt="" border="0"></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><table width="100%" bgcolor="#D6DBE4" border="0" cellspacing="0" cellpadding="6">
  		<tr>
    		<td><table width="100%" bgcolor="#FFFFFF"  border="0" cellspacing="2" cellpadding="4">
  				<tr><!-- LEFT COLUMN -->
    				<td width="68%" align="left" valign="top">
<%					' If ANY of the form inputs are blank, redisplay the form
					' This checks to see if the form has been completed
					
					If (Request.Form("name") = "" OR Request.Form("email") = "" OR Request.Form("subject") = "" OR Request.Form("message") = "") Then
					
						' Since AT LEAST ONE of the form inputs are blank, 
						' display Error Message because we want ALL of the fields completed
%>
					  	<p><img src="images/text_contact_the_artist.gif" width="410" height="40" border="0" alt="Contact the Artist"></p>
<%					  	If NOT Request.Form("status") = "1" Then %>
						  	<table align="center"  border="0" cellspacing="2" cellpadding="2">
						  	<tr>
						  		<td colspan="2">Contact me by  using one of the following methods:<br></td>
								</tr>
							<tr>
								<td width="45" height="47" valign="top"><img src="images/text_email.gif" width="40" height="40" border="0" alt="E-mail the Artist"></td>
								<td width="262" valign="bottom">
<!-- -------- start email address -------- -->
<SCRIPT type="text/javascript" language="javascript">
<!-- Begin
a="gerald";
b="geraldsmithartstudio.com";
document.write('<a href=\"mailto:'+a+'@'+b+'\">');
document.write(a+'@'+b+'<\/a>');
// End -->
</SCRIPT>
<!-- ----------- end email address ------------ --></td>
								</tr>
					  		</table>
					  
					  		<p align="center"><em><strong>or complete <br>and send the following form:</strong></em></p>
<%						End If
						If Request.Form("status") = "1" Then %>
							<BR>
							<div align="center"><STRONG>Error: Please enter ALL fields.</STRONG>
<%							' Initialize Error Message
							strError = "<BR>"
							
							If Request.Form("name") = "" Then
								strError = strError & "-- Name not supplied --<BR>"
							End If
							If Request.Form("email") = "" Then
								strError = strError & "-- E-mail not supplied --<BR>"
							End If
							If Request.Form("subject") = "" Then
								strError = strError & "-- Subject not supplied --<BR>"
							End If
							If Request.Form("message") = "" Then
								strError = strError & "-- Message not supplied --<BR>"
							End If
%>
							<%=strError%>
							</div>
<%						End If %>
		
						<!-- Display input form -->
					  	<form action="<%= Request.ServerVariables("SCRIPT_NAME") %>" method="post">
					  	<table align="center" width="322" bgcolor="#FFFFFF" cellpadding="0" cellspacing="1" border="0">
					  	<tr>
					  		<td width="53" align="right" bgcolor="#FFFFFF">From:</td>
							<td width="266" align="left" bgcolor="#FFFFFF"><input name="name" type="text" size="30" maxlength="60" value="<%= Request.Form("name") %>"></td>
							</tr>
					  	<tr>
					  		<td width="53" align="right" bgcolor="#FFFFFF">E-mail:</td>
							<td width="266" align="left" bgcolor="#FFFFFF"><input name="email" type="text" size="30" maxlength="100" value="<%= Request.Form("email") %>"></td>
							</tr>
					  	<tr>
					  		<td width="53" align="right" bgcolor="#FFFFFF">Subject:</td>
							<td width="266" align="left" bgcolor="#FFFFFF"><input name="subject" type="text" size="30" maxlength="50" value="<%= Request.Form("subject") %>"></td>
							</tr>
					  	<tr>
					  		<td bgcolor="#FFFFFF" align="left" colspan="2">Message:<br>
								<textarea name="message" cols="50" rows="5"><%= Request.Form("message") %></textarea></td>
							</tr>
					  	</table>
					  	<table align="center" width="322" cellpadding="0" cellspacing="1" border="0">
					  	<tr>
					  	<tr>
							<td align="right"><INPUT TYPE="Hidden" NAME="status" VALUE="1"><input name="sendform" type="Submit" class="formbutton" value="Send">&nbsp;&nbsp;<input type="Reset" class="formbutton" value="Reset"</td>
							</tr>
					  	</table>
						</form>
				
		<% 			Else 
						' This portion is executed only once user has 
						' completed all of the fields
						
						Dim strName			'Holds person's name or email
						strName = ""
						
						If Request.Form("name") <> "" Then
							strName = Trim(Request.Form("name"))
						Else
							strName = Request.Form("email")
						End If
						
						Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
						
						Mailer.FromName = strName
						Mailer.FromAddress = Trim(Request.Form("email"))
						Mailer.RemoteHost = "smtp.geraldsmithartstudio.com"
						Mailer.AddRecipient "Gerald Smith Art Studio", "gerald@geraldsmithartstudio.com"
						Mailer.Subject = Request.Form("subject")
						
						strMsgHeader = "MESSAGE FROM WEBSITE  O N L I N E  FORM" & _
						vbCrLf & vbCrLf
						strMsgContents = "From: " & Request.Form("name") & vbCrLf & _
						"Email: " & Request.Form("email") & vbCrLf & _
						"====================================================" & vbCrLf & vbCrLf & _
						Request.Form("message") & vbCrLf
						
						Mailer.BodyText = strMsgHeader & strMsgContents
						' Did E-Mail get sent?
						If Mailer.SendMail Then
							'Yes %>
							<BR>
							<h3 align="center">Thanks!<br>
							Your message has been sent to<br>
							Gerald at Gerald Smith Art Studio.</h3>
<%						else
							'No
							Response.Write "<BR><BR><BR>Sorry! Your message was not sent. Mail send failure. Error was " & Mailer.Response
						end if 
						
					End If %>
				</td>
					<!-- RIGHT COLUMN -->
    			<td width="32%" bgcolor="#D2E9F7" align="left" valign="top">
					<h4>Location of Art Studio </h4>
					<p>The art studio of <em>Gerald Steadman Smith</em> is located in the city of Ottawa (Stittsville), Ontario.</p>
					
					<h4>Art Studio Viewing</h4>
					<p>Personal viewing by appointment.</p>
					
					<h4>Commissions</h4>
					<p>Gerald Smith accepts proposals for commission.</p>
				  </td></tr>
				</table>				
				<p align="center"><SCRIPT LANGUAGE="JavaScript">
// Original: Kenneth Preston <drkennan@ionet.net> -->

// <!-- Begin
var m = "Page updated " + document.lastModified;
var p = m.length-8;
document.write(m.substring(p, 0));
document.writeln("<br><br>");
document.write('Copyright &copy; 2005');
if ((new Date()).getFullYear() > 2005) document.write('-' + (new Date()).getFullYear());
document.write('.');
// End -->
</script>
    Gerald Steadman Smith.  All rights reserved.<br>
    Content on this website is not to be reproduced without prior permission.
    <br/>
    <p id="hostedBy">
        Maintained by <a href="http://www.steven-smith.net/" target="_blank">Steve Smith</a>
    </p>
			</td>
  			</tr>
		</table>
	</td>
  </tr>
</table>
</div>
</body>
</html>
