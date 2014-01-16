<%@ Page Language="VB" debug=true  %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<!--#include file="../purchasing/header.aspx"-->
<!--#include file="../purchasing/functions.aspx"-->
<script language="vb" runat="server">
    Dim dbread As OleDbDataReader
     Dim objCommand As OleDbCommand 
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not Page.IsPostBack Then
            If Session("sifgecicino") Is Nothing Then
                Response.Redirect("sifgiris.aspx")
            Else       
                Data_bind()
            End If
        End If

    End Sub
    
    Sub Data_bind()
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()        
       
                  
        strSQL = "SELECT * FROM tbl_sifdetay_gecici where sifnodetay=" & Session("sifgecicino") & ""
        objCommand = New OleDbCommand(strSQL, objConn)
        dbread  = objCommand.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        DataList1.DataSource = dbread
        DataList1.DataBind()
        dbread.Close()
                      
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
        
    End Sub
    Protected Sub Update_data(ByVal sender As Object, ByVal e As CommandEventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        Dim id As Integer = e.CommandArgument
        strSQL = "UPDATE [tbl_sifdetay_gecici] SET [malzemekod] ='" & Trim(DropDownStok.SelectedItem.Value) & "', " & _
                       "[talepmiktar] = '" & Trim(txttalepmiktar.Text) & "' ,[birim] ='" & Trim(DropDownBirim.SelectedItem.Value) & "',[costcenter] ='" & Trim(DropDownCs.SelectedItem.Value) & "', " & _
                       "[cihazkod] = '" & Trim(DropDownMch.SelectedItem.Value) & "',[malzemetanim] = '" & Trim(txtmalzemeack.text) & "' ,[acilmi] ='" & Trim(DropDownAcilmi.SelectedItem.Text) & "' " & _
               "WHERE [ID] =" & id
        
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()
        Data_bind()         
    End Sub
    Sub delete_data(ByVal sender As Object, ByVal e As CommandEventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "Delete from tbl_sifdetay_gecici where ID = " & e.CommandArgument & ""
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()
        Data_bind()
    End Sub

    Protected Sub insert_data(ByVal sender As Object, ByVal e As System.EventArgs)
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
             
        strSQL = ""
        strSQL = strSQL & "INSERT INTO tbl_sifdetay_gecici "
        strSQL = strSQL & "(sifnodetay,createdate,malzemekod,malzemetanim,talepmiktar,birim,cihazkod,costcenter,acilmi) " & vbCrLf
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & "'" & Trim(Session("sifgecicino")) & "',"
        strSQL = strSQL & "'" & Trim(System.DateTime.Now.ToShortDateString()) & "',"
        strSQL = strSQL & "'" & Trim(DropDownStok.SelectedItem.Value) & "',"
        strSQL = strSQL & "'" & Trim(txtmalzemeack.Text) & "',"
        strSQL = strSQL & "'" & Trim(txttalepmiktar.Text) & "',"
        strSQL = strSQL & "'" & Trim(DropDownBirim.SelectedItem.Value) & "',"
        strSQL = strSQL & "'" & Trim(DropDownMch.SelectedItem.Value) & "',"
        strSQL = strSQL & "'" & Trim(DropDownCs.SelectedItem.Value) & "',"
        strSQL = strSQL & "'" & Trim(DropDownAcilmi.SelectedItem.Value) & "'"
        strSQL = strSQL & ");"
       
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        objCommand.Connection.Close()
        Data_bind()
    End Sub
    
     Protected Sub create_sif(ByVal sender As Object, ByVal e As System.EventArgs) 
     objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
     objConn.Open()
          
      strSQL = "SELECT * FROM tbl_sifdetay_gecici where sifnodetay=" & Session("sifgecicino") & ""
        objCommand = New OleDbCommand(strSQL, objConn)
        dbread  = objCommand.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
      If (dbread.HasRows()) Then         
        dim maxno as integer
        strSQL = "SELECT max(sifno) as maxsifno FROM tbl_sifmain"
        objCommand= New OleDbCommand(strSQL, objConn)
        Dim objDR As OleDbDataReader
        objDR = objCommand.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
        objDR.Read()
        If Not IsDBNull(objDR("maxsifno")) Then
           maxno=CInt(objDR("maxsifno"))                
        else
           maxno=0
        end if          
               
        strSQL = "INSERT INTO tbl_sifmain SELECT * FROM tbl_sifmain_gecici where sifno=" & Session("sifgecicino") & ""
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()
        strSQL = "INSERT INTO tbl_sifdetay SELECT * FROM tbl_sifdetay_gecici where sifnodetay=" & Session("sifgecicino") & ""
        objCommand = New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()        
        strSQL = "UPDATE [tbl_sifmain] SET [sifno] ='" & maxno + 1 & "',[onaydurum]=" & Session("userlevel") & " WHERE [sifno] =" & Session("sifgecicino")
        objCommand= New OleDbCommand(strSQL, objConn)
        objCommand.ExecuteNonQuery()      
        objCommand.Connection.Close()        
        response.redirect("main.aspx?sifno=" & maxno + 1)
        else
          response.redirect("sifgiris.aspx")
        End If
    End Sub
    
     
 
</script>	
<HTML>
<HEAD>
<title>database</title>
        <style type="text/css">
            .style1
            {
                font-family: Tahoma;
                text-align: right;
            }
        </style>
</HEAD>
<body bgcolor="#EFEFEF">
  
 <form id="form1" method="post" runat="server">
   
<%   
objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
objConn.Open()
    Dim onaydurum As String
        Dim sifdurum As String
    strSQL = "SELECT * FROM tbl_sifmain_gecici where sifno=" & Session("sifgecicino") & ""
    Dim objCmd As New OleDbCommand(strSQL, objConn)
    Dim objDR As OleDbDataReader
    objDR = objCmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)
    objDR.Read()
    If objDR("statusmain") = 0 Then
        sifdurum = "OPEN"
    Else
        sifdurum = "CLOSED"
    End If
    If objDR("onaydurum") = 0 Then
        onaydurum = "WAITING FOR APPROVE"
    ElseIf objDR("onaydurum") = 1 Then
        onaydurum = "DEPARTMENT MANAGER APPROVED"
    ElseIf objDR("onaydurum") = 2 Then
        onaydurum = "PURCHASING"
    End If
%>
<table align="center" bgcolor="Gainsboro"  width="80%" 
     style="font-family: Tahoma; font-size: xx-small">
    <tr>
        <td>
            <table align="center"  border="1" width="100%" bgcolor="#336699">
                <tr>
                    <td style="font-size: x-small; font-weight: 700 ; color:White" >
                        PO NUMBER:</td>
                    <td style="font-size: x-small; font-weight: 700; color:White">
                        PO DATE:</td>
                    <td style="font-size: x-small; font-weight: 700; color:White">
                        DEPARTMENT:</td>
                    <td style="font-size: x-small; font-weight: 700; color:White">
                        USER:</td>
                </tr>
                <tr>
                    <td bgcolor="#003366">
        <font color="White" size="2"                             
        style="font-size: x-small; color: #FFFFFF; font-weight: bold;" >
                      <%=Session("sifgecicino")%></font></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF" >
                        <%=FormatDateTime(objDR("createdate"),2)%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF" >
                        <%=departmentad(objDR("department"))%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=kisiad(objDR("username"))%></td>
                </tr>
                <tr>
                    <td style="font-size: x-small; font-weight: 700 ; color:White">
                        STATUS:</td>
                    <td style="font-size: x-small; font-weight: 700; color:White">
                        REQUIRED DATE:</td>
                    <td style="font-size: x-small; font-weight: 700; color:White">
                        REQUIRED COMPANY:</td>
                    <td style="font-size: x-small; font-weight: 700; color:White">
                        REF NO:</td>
                </tr>
                <tr>
                    <td bgcolor="#003366">
        <font color="White" size="2"                             
        style="font-size: x-small; color: #FFFFFF; font-weight: bold;" >
                    <%=sifdurum%></font></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=FormatDateTime(objDR("taleptarih"),2)%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF">
                        <%=firmaad(objDR("tercihfirma"))%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF" >
                       <%=objDR("talepno")%></td>
                </tr>
            </table>
        </td>
    </tr>
    <tr>
        <td>
       
          
            <table width="100%" >
             <asp:Repeater ID="DataList1" runat="server" >
                  <HeaderTemplate>               
                    <tr bgcolor="#003366">
                        <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Material</th>
                        <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Explanation</th>
                        <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Quantity</th>
                        <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Unit</th>
                        <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Cost Center</th>                       
                         <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Urgent?</th>
                        <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Device Code</th>
                         <th style="font-size: x-small; font-weight: 700; color: #FFFFFF;">
                            Process</th>
                    </tr>              
                    </HeaderTemplate>              
                    <ItemTemplate>             
                    <tr bgcolor="#99CCFF" style="font-family: Tahoma; font-size: xx-small; color: #000000">
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
                    <th>
                    <asp:LinkButton id="LinkButton1" 
                       Text="Delete"
                       CommandName="Order" 
                       CommandArgument=<%#Container.DataItem("ID")%>  
                       OnCommand="delete_data" 
                       runat="server"/>
                       <asp:LinkButton id="LinkButton2" 
                       Text="Update"
                       CommandName="Order" 
                       CommandArgument=<%#Container.DataItem("ID")%>  
                       OnCommand="Update_data" 
                       runat="server"/>                      
                   </th>
                </tr>               
            </ItemTemplate>
              </asp:Repeater>
               <tr bgcolor="#006666">
                    <td class="style1">                  
                        <asp:DropDownList ID="DropDownStok" runat="server"  ></asp:DropDownList>
                    </td>
                    <td class="style1">
                        <asp:TextBox ID="txtmalzemeack" runat="server" ></asp:TextBox>
                    </td>
                    <td class="style1">
                        <asp:TextBox ID="txttalepmiktar" onkeypress="onlyDigits(this,event)" 
                            runat="server" >0</asp:TextBox>
                    </td>
                    <td class="style1">
                        <asp:DropDownList ID="DropDownBirim" runat="server" >
                        </asp:DropDownList>
                    </td>
                     <td class="style1" >                  
                        <asp:DropDownList ID="DropDownCs" runat="server" ></asp:DropDownList>
                    </td>
                    <td class="style1"><asp:DropDownList id="DropDownAcilmi" runat="server" >
                            <asp:ListItem>E</asp:ListItem>
                            <asp:ListItem>H</asp:ListItem>                            
                        </asp:DropDownList>                        
                    </td>
                    <td class="style1">
                     <asp:DropDownList ID="DropDownMch" runat="server"  ></asp:DropDownList>
                        </td>
                    <td class="style1">
                        <asp:Button ID="Button1" runat="server" onclick="insert_data" Text="ADD" 
                            Height="20px" style="font-weight: 700"  />                        
                    </td>
                </tr>  
            </table>
       
        </td>
    </tr>
    
    <tr>
        <td>
            <table  border="1"  width="100%" bgcolor="#336699">
                <tr>
                    <td nowrap="nowrap" style="font-size: x-small; font-weight: 700 ;color: #FFFFFF">
                        Explanation:</td>
                    <td nowrap="nowrap" style="font-size: x-small; font-weight: 700 ;color: #FFFFFF">
                        Approve Status:</td>
                    <td nowrap="nowrap" style="font-size: x-small; font-weight: 700 ;color: #FFFFFF">
                    İşlem:</td>
                </tr>
                <tr>
                    <td style="font-size: xx-small; background-color: #FFFFFF ">
                        <%=objDR("aciklama")%></td>
                    <td style="font-size: xx-small; background-color: #FFFFFF" >
                        <%=onaydurum%></td>
                    <td style="font-size: xx-small; background-color: #003366; text-align: right;">
                            <asp:Button ID="KAYDET" runat="server" Text="REGISTER" onclick="create_sif" 
                                Height="21px" style="font-weight: 700; text-align: center;" Width="97px" />
                    </td>
                </tr>
               

            </table>
        </td>
    </tr>
</table>
<%  objConn.Close()%>
</form>
</body>
</HTML>
