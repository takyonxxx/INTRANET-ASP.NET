<%@ Page Language="VB" debug=true  %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<!--#include file="../purchasing/header.aspx"-->
<!--#include file="../purchasing/functions.aspx"-->
<script language="vb" runat="server">
dim duzeltstatus as integer
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not Page.IsPostBack Then
            If Session("username") Is Nothing Then
                Response.Redirect("../index.aspx")
            Else
                If (Request.QueryString("sifno")) Is Nothing Then
                    Response.Redirect("../index.aspx")
                Else
                    Session("sifgecicino") = Request.QueryString("sifno")
                    DataBind()
                End If
            End If
        End If
    End Sub
    
    Sub DataBind()
                objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                    objConn.Open()
                    strSQL = "SELECT * FROM tbl_sifdetay where sifnodetay=" & Session("sifgecicino") & ""
                    Dim dbcomm As OleDbCommand = New OleDbCommand(strSQL, objConn)
                    Dim dbread As OleDbDataReader = dbcomm.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
                    DataList1.DataSource = dbread
                    DataList1.DataBind()
                    dbread.Close()
                    objConn.Close()
     End Sub
     
     Protected Sub Update_data(ByVal sender As Object, ByVal e As CommandEventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        Dim id As Integer = e.CommandArgument
        strSQL = "UPDATE [tbl_sifdetay] SET [malzemekod] ='" & Trim(DropDownStok.SelectedItem.Value) & "', " & _
                       "[talepmiktar] = '" & Trim(txttalepmiktar.Text) & "' ,[birim] ='" & Trim(DropDownBirim.SelectedItem.Value) & "',[costcenter] ='" & Trim(DropDownCs.SelectedItem.Value) & "', " & _
                       "[cihazkod] = '" & Trim(DropDownMch.SelectedItem.Value) & "',[malzemetanim] = '" & Trim(txtmalzemeack.text) & "' ,[acilmi] ='" & Trim(DropDownAcilmi.SelectedItem.Text) & "' " & _
               "WHERE [ID] =" & id
        
        Dim objCommand As OleDbCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close() 
        objConn.Close()      
        DataBind()    
    End Sub
     Sub delete_data(ByVal sender As Object, ByVal e As CommandEventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "Delete from tbl_sifdetay where ID = " & e.CommandArgument & ""
        Dim objCommand As OleDbCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()  
        objConn.Close() 
        DataBind()    
    End Sub

    Sub btnOnayla_OnClick(ByVal Src As Object, ByVal E As EventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        Dim id As Integer = Request.QueryString("sifno")
        strSQL = "UPDATE [tbl_sifmain] SET [onaydurum] ='" & Session("userlevel") & "' " & _
                 "WHERE [sifno] =" & Session("sifgecicino")
        Dim objCommand As OleDbCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()
        objConn.Close()
        Response.Redirect("main.aspx?onaylayacaklarim=1")
    End Sub
    Sub btnRed_OnClick(ByVal Src As Object, ByVal E As EventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        Dim id As Integer = Request.QueryString("sifno")
        strSQL = "UPDATE [tbl_sifmain] SET [red] =1 " & _
                 "WHERE [sifno] =" & Session("sifgecicino")
        Dim objCommand As OleDbCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()
        objConn.Close()
        Response.Redirect("main.aspx?onaylayacaklarim=1")
    End Sub
    Sub btnDuzelt_OnClick(ByVal Src As Object, ByVal E As EventArgs)
        duzeltstatus =1  
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()                      
        Dim ds As DataSet
        Dim myda As OleDbDataAdapter = New OleDbDataAdapter("Select * from tbl_stok ", objConn)
        ds = New DataSet
        myda.Fill(ds, "AllTables")
        DropDownStok.DataSource = ds
        DropDownStok.DataSource = ds.Tables(0)
        DropDownStok.DataTextField = trim(ds.Tables(0).Columns("name").ColumnName.ToString())
        DropDownStok.DataValueField = trim(ds.Tables(0).Columns("code").ColumnName.ToString())
        DropDownStok.DataBind()
                       
        myda = New OleDbDataAdapter("Select * from tbl_birim ", objConn)
        ds = New DataSet
        myda.Fill(ds, "AllTables")
        DropDownBirim.DataSource = ds
        DropDownBirim.DataSource = ds.Tables(0)
        DropDownBirim.DataTextField = trim(ds.Tables(0).Columns("name").ColumnName.ToString())
        DropDownBirim.DataValueField = trim(ds.Tables(0).Columns("code").ColumnName.ToString())
        DropDownBirim.DataBind()
        
        myda = New OleDbDataAdapter("Select * from tbl_costcenter ", objConn)
        ds = New DataSet
        myda.Fill(ds, "AllTables")
        DropDownCs.DataSource = ds
        DropDownCs.DataSource = ds.Tables(0)
        DropDownCs.DataTextField = trim(ds.Tables(0).Columns("name").ColumnName.ToString())
        DropDownCs.DataValueField = trim(ds.Tables(0).Columns("code").ColumnName.ToString())
        DropDownCs.DataBind()
       
        myda = New OleDbDataAdapter("Select * from tbl_machine ", objConn)
        ds = New DataSet
        myda.Fill(ds, "AllTables")
        DropDownMch.DataSource = ds
        DropDownMch.DataSource = ds.Tables(0)
        DropDownMch.DataTextField = trim(ds.Tables(0).Columns("name").ColumnName.ToString())
        DropDownMch.DataValueField = trim(ds.Tables(0).Columns("code").ColumnName.ToString())
        DropDownMch.DataBind()         
        objConn.close()
        DataBind()
    End Sub
</script>	
<HTML>
<HEAD>
<title>SEE PO</title>
</HEAD>
<body bgcolor="#EFEFEF">
  
 <form id="form1" method="post" runat="server">   
<%   
objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
objConn.Open()
    Dim onaydurum As String
    Dim sifdurum As String
    Dim sifred As String
    strSQL = "SELECT * FROM tbl_sifmain where sifno=" & Request.QueryString("sifno") & ""
    Dim objCmd As New OleDbCommand(strSQL, objConn)
    Dim objDR As OleDbDataReader
    objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
    objDR.Read()
    If objDR("statusmain") = 0 Then
        sifdurum = "AÇIK"
    Else
        sifdurum = "KAPALI"
    End If
   If objDR("red") = 0 Then
        sifdurum = "APPROVED"
   Else
        sifdurum = "REFUSED"
   End If   
    
%>
<table align="center" bgcolor="Gainsboro" width="90%" 
     style="font-family: Tahoma; font-size: xx-small" >
    <tr>
        <td>
            <table align="center"  border="1" width="100%">
                <tr>
                    <td style="font-size: xx-small; font-weight: 700" >
                        PO NUMBER:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        PO DATE:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        DEPARTMENT:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        USER:</td>
                </tr>
                <tr>
       <td bgcolor="#003366">
        <font color="White" size="2"                             
        style="font-family: Tahoma; font-size: x-small; color: #FFFFFF; font-weight: bold;" >
                      <%=Session("sifgecicino")%></font></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF" >
                        <%=FormatDateTime(objDR("createdate"),2)%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=departmentad(objDR("department"))%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=kisiad(objDR("username"))%></td>
                </tr>
                <tr>
                    <td style="font-size: xx-small; font-weight: 700">
                        STATUS:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        REQUIRED DATE:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        REQUIRED COMPANY:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        REF NO:</td>
                </tr>
                <tr>
                    <td bgcolor="#003366">
        <font color="White" size="2"                             
        style="font-family: Tahoma; font-size: x-small; color: #FFFFFF; font-weight: bold;" >
                    <%=sifdurum%></font></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=FormatDateTime(objDR("taleptarih"),2)%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=firmaad(objDR("tercihfirma"))%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                       <%=objDR("talepno")%></td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>          
            <table width="100%>
             <asp:Repeater ID="DataList1" runat="server">
                  <HeaderTemplate>               
                    <tr bgcolor="#b0c4de">
                        <th style="font-size: xx-small; font-weight: 700">
                            Material</th>
                        <th style="font-size: xx-small; font-weight: 700">
                            Explanation</th>
                        <th style="font-size: xx-small; font-weight: 700">
                            Quantity</th>
                        <th style="font-size: xx-small; font-weight: 700">
                            Unit</th>
                        <th style="font-size: xx-small; font-weight: 700">
                            Cost Center</th>                        
                        <th style="font-size: xx-small; font-weight: 700">
                           Urgent?</th>
                        <th style="font-size: xx-small; font-weight: 700">
                           Device Number</th>               
                    </tr>              
                    </HeaderTemplate>              
                    <ItemTemplate>             
                    <tr bgcolor="#f0f0f0" style="font-family: Tahoma; font-size: xx-small; color: #000000">
                   <th><%#stokad(Container.DataItem("malzemekod"))%> 
                    </th>
                    <th><%#Container.DataItem("malzemetanim")%> 
                    </th>
                    <th>
                        <%#Container.DataItem("talepmiktar")%> 
                    </th>
                    <th>
                        <%#birimad(Container.DataItem("birim"))%> 
                    </th>
                    <th>
                        <%#csad(Container.DataItem("costcenter"))%> 
                    </th>                    
                    <th>
                        <%#Container.DataItem("acilmi")%> 
                    </th>
                    <th>
                        <%#mchad(Container.DataItem("cihazkod"))%> 
                    </th>
                    <% if duzeltstatus=1 then%>
                    <th>
                    <asp:LinkButton id="LinkButton1" 
                       Text="Sil"
                       CommandName="Order" 
                       CommandArgument=<%#Container.DataItem("ID")%>  
                       OnCommand="delete_data" 
                       runat="server"/>
                       <asp:LinkButton id="LinkButton2" 
                       Text="Düzelt"
                       CommandName="Order" 
                       CommandArgument=<%#Container.DataItem("ID")%>  
                       OnCommand="Update_data" 
                       runat="server"/>                      
                   </th>
                   <%end if%>
                </tr>                  
            </ItemTemplate>
              </asp:Repeater>
            </table>
       
        </td>
    </tr>
    <% if duzeltstatus=1 then%>
     <tr>
        <td> 
           <table align="center"  width="100%">
             <tr>
                    <td>                  
                        <asp:DropDownList ID="DropDownStok" runat="server" Width="151px" Height="16px"></asp:DropDownList>
                    </td>
                    <td>
                        <asp:TextBox ID="txtmalzemeack" runat="server" Width="68px"></asp:TextBox>
                    </td>
                    <td>
                        <asp:TextBox ID="txttalepmiktar" onkeypress="onlyDigits(this,event)" 
                            runat="server" Width="68px">0</asp:TextBox>
                    </td>
                    <td>
                        <asp:DropDownList ID="DropDownBirim" runat="server">
                        </asp:DropDownList>
                    </td>
                     <td >                  
                        <asp:DropDownList ID="DropDownCs" runat="server" Width="142px" Height="17px"></asp:DropDownList>
                    </td>
                    <td><asp:DropDownList id="DropDownAcilmi" runat="server">
                            <asp:ListItem>E</asp:ListItem>
                            <asp:ListItem>H</asp:ListItem>                            
                        </asp:DropDownList>                        
                    </td>
                    <td>
                     <asp:DropDownList ID="DropDownMch" runat="server" Width="142px" Height="17px"></asp:DropDownList>
                        </td>
                    <td>      
                        
                    </td>
                </tr>
            </table>
        </td>
    </tr>
    <%end if%>
    <tr>
        <td>
            <table  width="100%">
                <tr>
                    <td style="font-size: xx-small; font-weight: 700">
                        Explanation:</td>
                    <td style="font-size: xx-small; font-weight: 700">
                        Approve Status:</td>    
                        <%If (Request.QueryString("yoneticionay")) IsNot Nothing Then%>  
                         <td ></td>                              
                        <%End If%>               
                </tr>
                <tr>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=objDR("aciklama")%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%If objDR("red") = 0 Then
                        response.write (onayad(objDR("onaydurum")))
                        ELSE
                        response.write(sifdurum)
                        END IF
                        %></td>  
                        <%If (Request.QueryString("yoneticionay")) IsNot Nothing Then%>  
                         <td >
                             <asp:Button ID="onayla" runat="server" Text="ONAYLA" OnClick="btnOnayla_OnClick" />
                             <asp:Button ID="duzelt" runat="server" Text="DÜZELT" OnClick="btnDuzelt_OnClick" />
                             <asp:Button ID="reddet" runat="server" Text="REDDET" OnClick="btnRed_OnClick"/></td>                              
                        <%End If%>                       
                </tr>
             </table>
        </td>
    </tr>
</table>
<%  objConn.Close()%>
</form>
</body>
</HTML>
