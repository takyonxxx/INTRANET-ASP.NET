<%@ Page Language="VB" AutoEventWireup="false" CodeFile="login.aspx.vb" Inherits="login" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%  
    If Request.QueryString("status") = "logoff" Then
        Session("username") = ""
    End If
    If Session("username") <> "" Then
        Response.Redirect("index.aspx")
    End If
    If Request.QueryString("enter") <> "" Then
        If Request.Form("username") <> "" And Request.Form("password") <> "" Then
            Dim objConnection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("intranet_data.mdb") & ";")
            Dim objCommand1 As OleDbCommand
            Dim tbluser As OleDbDataReader
            Dim strSQLQueryUser As String
            objConnection.Open()
            strSQLQueryUser = "SELECT * FROM tbl_user where user_name='" & Request.Form("username") & "' AND password='" & Request.Form("password") & "'"
            objCommand1 = New OleDbCommand(strSQLQueryUser, objConnection)
            tbluser = objCommand1.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
            If tbluser.HasRows() Then
                While tbluser.Read()
                    Session("username") = tbluser("user_name")
                End While
                Response.Redirect("index.aspx")
            Else
                Response.Redirect("login.aspx?error=user")
            End If
        Else
            Response.Write("<center><b>Please fill all fields</b></center>")
        End If
    End If
    If Request.QueryString("error") = "user" Then
        Response.Write("<center><b>Your username or password is wrong. Please try again.</b></center>")
    End If

%>
<head runat="server">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-9">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1254">
    <title>URSA LOGIN</title>
</head>
<body bgcolor="#EFEFEF">
    <TABLE cellSpacing=3 cellPadding=3  align=center border=0 style="WIDTH: 25%" 
        bgColor   =gainsboro >
      <FORM action="login.aspx?enter=ok" method=post id="form2" name="form1" autocomplete="on">
       <tr align=middle bgcolor=OrangeRed>
      <td colspan=2 bgcolor="#000099">
       <FONT size=2 face=tahoma color=White><B>URSA ISI LOGIN PAGE</B></FONT>
      </td>
      </tr>
      <tr align=middle>
      <td colspan=2 bgcolor="<%=Session("company_color")%>">
        <A href="changepassword.aspx"><FONT size=2 face=tahoma color=White>Change Password</FONT></A>
      </td>
      </tr>
      <TR>
      <TD><FONT face=tahoma size=2 color=navy><STRONG>&nbsp;Username:</STRONG></FONT></TD>
      <TD><INPUT id=text1 name=username style="WIDTH:120px; HEIGHT: 19px" > *</TD>
      </TR>
      <TR>  
      <TD><FONT face=tahoma size=2 color=navy><STRONG>&nbsp;Password:</STRONG></FONT></TD>
      <TD><INPUT id=password1 type=password name=password style="WIDTH:120px; HEIGHT: 19px"> *</TD>
      </TR>
      <TR><TD colspan=2 align=center bgcolor="<%=Session("company_color")%>">&nbsp;
      <INPUT class=button id=submit1 name=submit1 
              style="WIDTH: 130px; HEIGHT: 22px; font-family: Tahoma; font-weight: bold;" 
              type=submit value=Giris>
      </TD></TR>
      </FORM>
    </TABLE>
</body>
</html>
