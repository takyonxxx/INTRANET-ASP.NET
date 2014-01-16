<%@ Page Language="VB" AutoEventWireup="false" CodeFile="index.aspx.vb" Inherits="Company_Anasayfa" Debug="true"%>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%  On Error Resume Next
    Dim objConnection As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("intranet_data.mdb") & ";")
    Dim objCommand1 As OleDbCommand
    Dim objCommand2 As OleDbCommand
    Dim objCommand3 As OleDbCommand
    Dim objCommand4 As OleDbCommand
    Dim objCommand5 As OleDbCommand
    Dim tbllink As OleDbDataReader
    Dim tbllinkdetay As OleDbDataReader
    Dim tbldoviz As OleDbDataReader
    Dim tbluser As OleDbDataReader
    Dim tblcompany As OleDbDataReader
    Dim strSQLQuerylink As String
    Dim strSQLQuerylinkdetay As String
    Dim strSQLQueryuser As String
    Dim strSQLQuerycompany As String
    Dim linkname As String
    Dim link As String
    Dim companyname As String
    Dim i As Integer
    objConnection.Open()
    If Session("username") = "" Then
        objCommand1 = New OleDbCommand("SELECT * FROM tbl_company", objConnection)
        tblcompany = objCommand1.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        While tblcompany.Read()
            Session("company_name") = tblcompany("name")
            Session("company_color") = tblcompany("color")
        End While
        Response.Redirect("Login.aspx")
    Else
        
        strSQLQuerylink = "SELECT * FROM tbl_link"
        objCommand2 = New OleDbCommand(strSQLQuerylink, objConnection)
        tbllink = objCommand2.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
       
        strSQLQuerylinkdetay = "SELECT * FROM tbl_link_detay"
        objCommand3 = New OleDbCommand(strSQLQuerylinkdetay, objConnection)
        tbllinkdetay = objCommand3.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        
        strSQLQueryuser = "SELECT * FROM tbl_user where user_name='" & Session("username") & "'"
        objCommand4 = New OleDbCommand(strSQLQueryuser, objConnection)
        tbluser = objCommand4.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        While tbluser.Read()
            Session("user_name") = tbluser("name")
            Session("user_surname") = tbluser("surname")
            Session("user_mail") = tbluser("mail")
            Session("userlevel") = tbluser("level")
            Session("userbolum") = tbluser("department")
        End While
        objCommand1 = New OleDbCommand("SELECT * FROM tbl_company", objConnection)
        tblcompany = objCommand1.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        While tblcompany.Read()
            Session("company_name") = tblcompany("name")
            Session("company_color") = tblcompany("color")
        End While
    End If
   
    
%> 
<head runat="server">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=iso-8859-9">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;charset=windows-1254">
<title><%=Session("company_name")%></title>
    <style type="text/css">
        .style1
        {
            width: 343px;
            height: 18px;
        }
        .style2
        {
            width: 285px;
        }
        .style4
        {
            font-weight: bold;
            font-family: Tahoma;
            font-size: xx-small;
        }
        .style5
        {
            font-family: Tahoma;
        }
        .style6
        {
            font-family: Tahoma;
            font-weight: bold;
        }
        .style7
        {
            width: 67px;
        }
        .style8
        {
            font-size: large;
            font-weight: bold;
        }
        .style9
        {
            text-decoration: underline;
        }
        .style10
        {
            font-family: Tahoma;
            width: 46px;
        }
        .style11
        {
            font-family: Tahoma;
            font-weight: bold;
            color: #003366;
        }
        .style12
        {
            color: #003399;
        }
        .style13
        {
            font-weight: bold;
            text-decoration: none;
        }
        </style>
</head>
<body bgcolor="#EFEFEF">

<table width="80%" ALIGN="center" BORDER="0" CELLSPACING="0" CELLPADDING="0" >
	<tr>
	<td  align="center" bgcolor="<%=Session("company_color")%>" class="style7"></td>	
	<td align="center" bgcolor="<%=Session("company_color")%>" class="style8">
		<font color=White style="text-align: left"><%=Session("company_name")%> INTRANET PORTAL</font></td>
	<td  align="center" bgcolor="<%=Session("company_color")%>" class="style7"></td>
	</tr>	
</table>

<table WIDTH="80%" ALIGN="center" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr><hr color=DarkGray>
		<td bgcolor=DimGray colspan="2">&nbsp;<img alt="" src="images/dir_plus.gif" style="width: 12px; height: 12px" />
    <A HREF="index.aspx"><font face=tahoma color=White size="2"><b>MAIN PAGE</b></font></A>
      </td>
	</tr>	
	<tr>
	<%If Request.QueryString("id") = "" Then%>
		<td align="left" valign="top" bgcolor=#DFDFDF class="style2"><hr />
		<%		    While tbllink.Read()
		        linkname = tbllink("linkname")
		        link = CStr(tbllink("link"))
		        ID = tbllink("id")
		%>&nbsp;
       <img src="images/orange_line.jpg" style="height: 6px; width: 16px" />&nbsp;
       <a HREF="<%=link%>?id=<%=id%>"><font face=tahoma color=black size="2"><b><%=linkname%></b></font></a><hr />
        <%  
        End While%>
		</td>
		<%End If%>
		
		<%if Request.QueryString ("id")<>"" then%> 
		<td align="left" valign="top" bgcolor=#DFDFDF>
		<%
		    While tbllink.Read()
		        If CInt(Request.QueryString("id")) = CInt(tbllink("id")) Then%>
		       &nbsp; <img src="images/orange_line.jpg" style="height: 6px; width: 16px" />&nbsp;
                  <A href="javascript:history.go(-1)">
                   <font face=tahoma color="Black" size="2"><b><%=tbllink("linkname")%></b></font></A><br/>
        <%    End If
		    End While
        i = 0
        While tbllinkdetay.Read()
		    If CInt(Request.QueryString("id")) = tbllinkdetay("id") Then
		        link = tbllinkdetay("link")
			        %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                <img src="images/ok.GIF" style="height: 8px; width: 16px" />&nbsp;&nbsp;<a HREF="<%=tbllinkdetay("link")%>"><font face=tahoma color=MidnightBlue size="2"><%=tbllinkdetay("linkname")%></font></a><br/>
			        <%i=i+1
			        end if			
                End While
            If i = 1 Then
			    Response.Redirect(link)
            End If%></td>
       <% End If%>	
    		
	</tr>
	
</table> 	

<table width="80%" align="center" border=0>
    <tr>  <hr color=DarkGray>
    <td align="left" nowrap valign=middle bgcolor="<%=Session("company_color")%>" class="style1">
        <span class="style4" style="color: #FFFF00">DATE:</span>&nbsp;<b><span style="color: #FFFFFF" class="style4"><%=FormatDateTime(Date.Today(), DateFormat.GeneralDate)%>
        &nbsp;&nbsp; </b>
        </span>
        <b>
        <span style="color: #FFFF00" class="style4">USER:</span><span style="color: #FFFFFF" class="style4">
        <span class="style4"> <%=Session("user_name") & " " & Session("user_surname")%></span></span></b>
        </td>
    </tr>
    <tr align=left bgcolor=DimGray><td>&nbsp;
    <a HREF="login.aspx?status=logoff" class="style13"><font face=tahoma color=white size="2">LOG OUT</font></a>
    </td></tr>
   </table>
<%  tbllink.Close()
    objConnection.Close()%>
</body>
</html>
</html>
