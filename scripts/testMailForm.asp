<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN">
<HTML>

<!-- --------------------------------------------------------------------------------------
     Alentus Default Web Site

       Page.......: ASP Mail Test Form
       Domain.....: geraldsmithartstudio.com
       URL........: http://www.geraldsmithartstudio.com/scripts/testMailForm.asp

     Copyright 2001, Alentus Corporation
     -------------------------------------------------------------------------------------- -->

<!-- HTML Header -->
<HEAD>
  <TITLE>Welcome to www.geraldsmithartstudio.com</TITLE>
  <META name='keywords' content='geraldsmithartstudio.com,www.geraldsmithartstudio.com'>
</HEAD>

<STYLE>
	<!-- body	{ font-family:arial;                                    } -->
	<!-- td		{ font-family:arial; font-size:10pt; font-weight:normal } -->
	<!-- th		{ font-family:arial; font-size:10pt; font-weight:bold   } -->
	<!-- p		{ font-family:arial; font-size:10pt;                    } -->
</STYLE>

<!-- Page Body -->
<BODY bgcolor=FFFFFF>

<!-- Page Banner -->
<HR><CENTER>
<FONT size=4>Welcome to www.geraldsmithartstudio.com</FONT>
<FONT size=2>
<HR>

<!-- Page Title -->
ASP Mail Test Form.
<P>
<IMG src='/images/ASPMailLogo.gif' align=center border=0>
<P>

<!-- ASP Mail Test Form -->
<FORM method=get action="testMailSend.asp">
<TABLE border="0">
  <TR>
    <TD>Enter Mail To Address:</TD>
    <TD><input type=text name="addressto" size=40></TD>
  </TR>
  <TR>
    <td valign="top">Enter Mail Message:</TD>
    <TD><textarea name="txtmsg" rows=3 cols=40>Hello!</textarea></TD>
  </TR>
  <TR>
    <td valign="top">&nbsp;</TD>
    <TD><input type="submit" value="Send Mail"></TD>
  </TR>
</TABLE>
</FORM>

<P>
<A href=/>Click here</A> to return to the home page.

<!-- Page Footer -->
<HR>
<P>www.geraldsmithartstudio.com is hosted by:
<BR>
<A href='http://www.alentus.com'><IMG src='/images/AlentusLogo.gif' align=center border=0></A>
</FONT>
</BODY>
</HTML>
