<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>

<%
    If Request.QueryString("change") <> "" Then
        If Request.Form("txt_username") <> "" And Request.Form("txt_oldpsw") <> "" Then
            Dim objConnection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("intranet_data.mdb") & ";")
            Dim objCommand1 As OleDbCommand
            Dim tbluser As OleDbDataReader
            Dim username_temp As String
            Dim strSQLQueryUser As String
            objConnection.Open()
            strSQLQueryUser = "SELECT * FROM tbl_user where user_name='" & Request.Form("txt_username") & "' AND password='" & Request.Form("txt_oldpsw") & "'"
            objCommand1 = New OleDbCommand(strSQLQueryUser, objConnection)
            tbluser = objCommand1.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
            If tbluser.HasRows() Then
                While tbluser.Read()
                    username_temp = tbluser("user_name")
                End While
                Dim strTxt1 As String = Request.Form("txt_newpsw1")
                Dim strTxt2 As String = Request.Form("txt_newpsw2")
                If strTxt1 <> "" And strTxt2 <> "" Then
                    If strTxt1 = strTxt2 Then
                        Dim strSQL As String = "UPDATE tbl_user SET tbl_user.password ='" & Trim(strTxt2) & "' " & _
                        "WHERE tbl_user.user_name = '" & username_temp & "' "
                        Dim myCommand As OleDbCommand = New OleDbCommand(strSQL, objConnection)
                        myCommand.ExecuteNonQuery()
                        Response.Redirect("login.aspx")
                    Else
                        Response.Write("<center><b>New password not match</b></center>")
                    End If
                Else
                    Response.Write("<center><b>New password can not be empty</b></center>")
                End If
               
            ElseIf Request.QueryString("hata") <> "user" Then
                Response.Write("<center><b>Your username or password is wrong. Please try again.</b></center>")
            End If
        End If
        If Request.Form("txt_username") = "" Or Request.Form("txt_oldpsw") = "" Then
            Response.Write("<center><b>Please fill the boxes</b></center>")
        End If
    End If
        
%>	
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-9">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1254">
</HEAD>
<BODY bgcolor="#EFEFEF">

<FORM action="changepassword.aspx?change=sifre" method=post id=form1 name=form1>

<TABLE cellSpacing=3 cellPadding=2 width="30%" align=center border=0 bgcolor=orangered style="WIDTH: 30%">
  
  <TR>
    <TD>
      <STRONG><center><FONT size=3 color=white>Change Your Password !</FONT></center></STRONG> 
    
      <TABLE cellSpacing=1 cellPadding=0 width="100%" align=center border=0 bgColor=gainsboro style="WIDTH: 95%">
        
        <TR><hr>
          <TD>&nbsp;<STRONG>Username</STRONG> </TD>
          <TD><INPUT id=text1 name=txt_username> *</TD></TR>
        <TR>
          <TD>&nbsp;<STRONG>Old Password</STRONG> </TD>
          <TD><INPUT id=password1 type=password name=txt_oldpsw> *</TD></TR>
        <TR>
          <TD>&nbsp;<STRONG>New Password</STRONG> </TD>
          <TD><INPUT id=password2 type=password name=txt_newpsw1> *</TD></TR>
        <TR>
          <TD>&nbsp;<STRONG>New Password Again</STRONG>  </TD>
          <TD><INPUT id=password3 type=password name=txt_newpsw2> *</TD></TR>
        <TR>
          <TD>
</TD>
          <TD>
            <P align=left><INPUT id=submit1 class="button" type=submit value="Set Password" name=change> 
            &nbsp;<A href="javascript:history.go(-1)">Back</A> 
        </P></TD></TR></TABLE></TD></TR></TABLE>
</FORM>

</BODY>
</HTML>
