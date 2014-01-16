<%@ Page Language="VB" debug=true  %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<!--#include file="../purchasing/functions.aspx"-->
<script language="vb" runat="server">
      
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not Page.IsPostBack Then
            If (Session("username") Is Nothing Or Session("userlevel") Is Nothing Or Session("userbolum") Is Nothing) Then
                Response.Redirect("../index.aspx")
            Else
                If (Request.QueryString("sifno")) IsNot Nothing Then
                    Session("sifno") = Request.QueryString("sifno")
                Else
                    Session("sifno") = Nothing
                End If
                objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                objConn.Open()
                Dim objCommand As OleDbCommand
                strSQL = "SELECT * FROM tbl_user where user_name='" & Session("username") & "'"
                objCommand = New OleDbCommand(strSQL, objConn)
                Dim dbread As OleDbDataReader = objCommand.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
                if (dbread.read())
                     session("usersatinalma")=dbread("satinalma")
                else
                session("usersatinalma")=""
                end if
                dbread.Close()
                If Request.QueryString("onayla") IsNot Nothing And Request.QueryString("sifno") IsNot Nothing Then  
                 objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                    objConn.Open()               
                    Dim id As Integer = Request.QueryString("sifno")
                    strSQL = "UPDATE [tbl_sifmain] SET [onaydurum] ='" & Session("userlevel") & "' " & _
                             "WHERE [sifno] =" & id
                    objCommand = New OleDbCommand(strSQL, objConn)
                    objCommand.ExecuteNonQuery()
                    objCommand.Connection.Close()
                ElseIf Request.QueryString("red") IsNot Nothing And Request.QueryString("sifno") IsNot Nothing Then
                 objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                    objConn.Open()
                    Dim id As Integer = Request.QueryString("sifno")
                    strSQL = "UPDATE [tbl_sifmain] SET [red] =1 " & _
                             "WHERE [sifno] =" & id
                    objCommand = New OleDbCommand(strSQL, objConn)
                    objCommand.ExecuteNonQuery()
                    objCommand.Connection.Close()
                End If
                objConn.Close()
        End If
        End If
    End Sub
</script>	
 <script language="javascript" type="text/javascript">

  function setLocation(href)
        {
        if (document.getElementById(href).value=="") 
           {
           location.href ='main.aspx';
           }
        else
           {
           location.href ='sifgor.aspx?sifno=' + document.getElementById(href).value;
           }
        }
  
 </script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <style type="text/css">
        .style1
        {
            width: 80%;  
            align="center"        
        }
        #Button2
        {
            height: 20px;
        }
        #Text1
        {
            width: 66px;
        }
        #sifno
        {
            width: 29px;
            height: 16px;
        }
        .style3
        {
            height: 25px;
            font-size: x-small;
            font-family: Tahoma;
        }
        #Text2
        {
            width: 65px;
        }
        #Button1
        {
            height: 20px;
        }
        .style6
        {
            font-family: Tahoma;
            font-size: x-small;
        }
        .style7
        {
            color: #FFFFFF;
        }
        .style10
        {
            background-color: #FFFFFF;
        }
        .style13
        {
            font-family: Tahoma;
            font-size: xx-small;
            font-weight: bold;
            width: 109px;
            background-color: #FFFFFF;
            height: 16px;
        }
        .style16
        {
            font-size: x-small;
        }
        .style17
        {
            font-size: x-small;
            color: #000000;
        }
        .style18
        {
            color: #003399;
        }
        .style20
        {
            text-decoration: none;
        }
    </style>
</head>
<body bgcolor="#EFEFEF" style="font-size: x-small">
    <form id="form1" runat="server">
    <table border="1" class="style1">
       
        <tr>
              <table width="80%" bgcolor="White" class="style1" cellpadding="1" cellspacing="1" align="center">
              <tr><td bgcolor="#006699" 
                      
                      style="color: #FFFFFF; font-weight: 700; font-family: Tahoma; font-size:small; background-color: #006699;">
                 PURCHASING SYSTEM -
              <a href="../index.aspx" 
                      style="font-family: Tahoma; font-size: x-small; font-weight: bold; color: yellow; text-decoration: none;">MAIN PAGE</a><hr 
                      style="height: 10px; background-color: #666666" />
                  </td>
    
        </tr>
                    <tr bgcolor="#006699">
                        <td style="font-weight: 700; background-color: #666666;" class="style7">
                            PROCESSES</td>
                    </tr>
                    
                    <tr>
                        
                        <td class="style6">
                            <a href="sifgiris.aspx" class="style20"><span class="style18">
                                <img src="images/bg.gif" style="height: 8px; width: 6px" /> CREATE PO</span></a></td>
                    </tr>
                    <tr>
                        <td colspan="0" class="style6">                           
                         <a href="main.aspx?onaylayacaklarim=1" class="style20"><span class="style18"> <img src="images/bg.gif" style="height: 8px; width: 6px" /> PENDING PO'S SHOULD APPROVE BY ME</span></a><span 
                                class="style6">
                            <table class="style1">  
                        <%  If Request.QueryString("onaylayacaklarim") = "1" Then%>
                         <tr><td class="style13" nowrap="nowrap">PO NUMBER / USER</td></tr>
                        <%
                                objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                                objConn.Open()
                                If Session("userlevel") = 3 Then
                                    strSQL = "SELECT * FROM tbl_sifmain WHERE onaydurum=" & CInt(Session("userlevel")) - 1 & " and red=0 "
                                Else
                                    strSQL = "SELECT * FROM tbl_sifmain WHERE onaydurum=" & CInt(Session("userlevel")) - 1 & " and department='" & Session("userbolum") & "' and red=0 "
                                End If
                                
                                
                                Dim objCmd As New OleDbCommand(strSQL, objConn)
                                Dim objDR As OleDbDataReader
                                objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
                                While (objDR.Read())
                                %>
                                <tr><td class="style10"><a href="sifgor.aspx?sifno=<%=objDR("sifno")%>&yoneticionay=1"
                                    style="font-family: Tahoma; font-size: xx-small; color: #FF0000; font-weight: bold; text-decoration: none;">
                                     <%=objDR("sifno")%> / <%=kisiad(objDR("username"))%></a> 
                                    <a href="main.aspx?sifno=<%=objDR("sifno")%>&onayla=1&onaylayacaklarim=1"
                                    style="font-family: Tahoma; text-decoration: none;"><span class="style17">APPROVE</span></a><span 
                                        class="style16">&nbsp; </span>
                                     <a href="main.aspx?sifno=<%=objDR("sifno")%>&red=1&onaylayacaklarim=1"
                                    style="font-family: Tahoma; text-decoration: none;"><span class="style17">REFUSE</span></a>
                                     </td>
                                </tr>                                
                                <%
                                End While
                            End If%>
                          </table>
                        </td>
                                        
                    </tr>
                    <tr>
                        <td class="style3" colspan="0">
                           <a href="main.aspx?onaylanacaklar=1" class="style20"><span class="style18"> <img src="images/bg.gif" style="height: 8px; width: 6px" /> PENDING PO'S SHOULD APPROVE BY MANAGER</span></a>
                            <table class="style1">                             
                          
                        <%  If Request.QueryString("onaylanacaklar") = "1" Then%>
                         <tr><td class="style13" nowrap="nowrap">PO NUMBER / USER</td></tr>
                        <%
                                objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                                objConn.Open()
                               
                                    strSQL = "SELECT * FROM tbl_sifmain WHERE onaydurum=" & CInt(Session("userlevel")) & " and username='" & Session("username") & "' and red=0 "
                                                              
                                Dim objCmd As New OleDbCommand(strSQL, objConn)
                                Dim objDR As OleDbDataReader
                                objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
                                While (objDR.Read())
                                %>
                                <tr><td class="style10"><a href="sifgor.aspx?sifno=<%=objDR("sifno")%>"
                                    style="font-family: Tahoma; font-size: xx-small; color: #FF0000; font-weight: bold; text-decoration: none;">
                                     <%=objDR("sifno")%> / <%=kisiad(objDR("username"))%></a>                                     
                                     </td>
                                </tr>                                
                                <%
                                End While
                            End If%>
                          </table>
                          </td>
                    </tr>                    
                </table>               
                      
        </tr>
         <tr>
             <table width="90%" bgcolor="White" class="style1" cellpadding="1" cellspacing="1" align="center">
                <tr bgcolor="#006699"><td style="font-weight: 700; background-color: #666666;" 
                        class="style7">REPORTS</td></tr>
                 <tr>
                        <td>
                            <input id="Text1" size="25" type="text" value='<%=Session("sifno")%>' />
                            <input id="Button2" onclick="setLocation('Text1');" type="button" 
                                value="Sif Gör" />
                            <hr />
                        </td>                            
                    </tr>   
                    <tr>
                     <td class="style6">
                             <a href="sifrapor.aspx" class="style20"><span class="style18"> <img src="images/bg.gif" style="height: 8px; width: 6px" /> SEARCH PO</span></a>
                             <hr />
                        </td>
                    </tr>               
                </table>
            </tr>
              <%  If session("usersatinalma") = "1" Then%>
            <tr>
             <table width="90%" bgcolor="White" class="style1" cellpadding="1" cellspacing="1" align="center">
                <tr bgcolor="#006699"><td style="font-weight: 700; background-color: #666666;" 
                        class="style7">PO PROCESSES</td></tr>
                 
                    <tr>
                        <td class="style6">
                             <a href="sifislem.aspx" class="style20"><span class="style18"> <img src="images/bg.gif" style="height: 8px; width: 6px" /> PO SEE-CLOSE</span></a><hr />
                        </td>
                    </tr>
                    
                </table>
            </tr>
            <%end if %>      
   
    </table>          
    </form>
</body>
</html>
