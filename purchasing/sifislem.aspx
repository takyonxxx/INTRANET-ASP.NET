<%@ Page Language="VB" debug="true" EnableViewState="False"%>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<!--#include file="../purchasing/header.aspx"-->
<!--#include file="../purchasing/functions.aspx"-->
<head>
    <style type="text/css">
        #fatno
        {
            width: 80px;
            height: 18px;
            font-size: xx-small;
            font-family: Tahoma;
        }
        #Text1
        {
            width: 62px;
            height: 18px;
            font-family: Tahoma;
            font-size: xx-small;
        }
        .style1
        {
            background-color: #FFFFFF;
        }
    </style>
</head>
<script language="vb" runat="server">
    Dim dbread, dbbirim,dbsifcontrol As OleDbDataReader
    Dim dbcomm As OleDbCommand
    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        If Not Page.IsPostBack Then
            If Session("username") Is Nothing And Session("usersatinalma") <> "1" Then
                Response.Redirect("../index.aspx")
            End If
            Data_bind()
        End If
        If Request.Form("sifkapat") <> "" Then
            objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
            objConn.Open()
            Dim ID As Integer = Request.QueryString("kapat")
            dim sifno as integer=cint(request.querystring("sifno"))
            strSQL = "UPDATE [tbl_sifdetay] SET [fatno] ='" & Request.Form("fatno") & "' ," & _
            "[onaymiktar] =" & CInt(Request.Form("teminmiktar")) & ", " & _
            "[birim] ='" & Cstr(Request.Form("birim")) & "', " & _
            "[statusdetay] =1 " & _
            "WHERE [ID] =" & ID
            Dim objCommand As OleDbCommand = New OleDbCommand(strSQL, objConn)
            objCommand.ExecuteNonQuery()          
            
            
            strSQL = "SELECT tbl_sifmain.*,tbl_sifdetay.statusdetay " & _
                     "FROM tbl_sifmain INNER JOIN tbl_sifdetay ON tbl_sifmain.sifno = tbl_sifdetay.sifnodetay " & _
                     "WHERE (statusdetay=0 AND sifnodetay= " & sifno & ") "
                    dbcomm = New OleDbCommand(strSQL, objConn)
                    dbsifcontrol = dbcomm.ExecuteReader()
            if not dbsifcontrol.hasrows() then
                    strSQL = "UPDATE [tbl_sifmain] SET [statusmain] =1,[onaydurum] =4 " & _
                             "WHERE [sifno] =" & sifno
                    objCommand = New OleDbCommand(strSQL, objConn)
                    objCommand.ExecuteNonQuery()        
            end if
            dbsifcontrol.close()            
            objCommand.Connection.Close()
            objConn.Close()
            Data_bind()
        End If
          
    End Sub
    Sub Data_bind()
        objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
        objConn.Open()
        strSQL = "SELECT tbl_sifmain.*,tbl_sifdetay.ID as sifdetayid, tbl_sifdetay.costcenter, tbl_sifdetay.malzemekod, tbl_sifdetay.malzemetanim, tbl_sifdetay.talepmiktar, tbl_sifdetay.birim, tbl_sifdetay.acilmi, tbl_sifdetay.cihazkod, tbl_sifdetay.createdate , tbl_sifdetay.fatno, tbl_sifdetay.onaymiktar,tbl_sifdetay.sifnodetay , tbl_sifdetay.createdate as sifdetaytarih " & _
                 "FROM tbl_sifmain INNER JOIN tbl_sifdetay ON tbl_sifmain.sifno = tbl_sifdetay.sifnodetay " & _
                 "WHERE (tbl_sifmain.onaydurum = 3 AND tbl_sifdetay.statusdetay=0) ORDER BY tbl_sifdetay.createdate ASC"
        dbcomm = New OleDbCommand(strSQL, objConn)
        dbread = dbcomm.ExecuteReader()
        strSQL = "SELECT * from tbl_birim"
        dbcomm = New OleDbCommand(strSQL, objConn)
        dbbirim = dbcomm.ExecuteReader()
         if not dbread.hasrows() then
         response.redirect("main.aspx?onaylayacaklarim=1")
         end if
 End Sub
</script>	
 <body bgcolor="#EFEFEF">      
<table align="center" bgcolor="Gainsboro" cellpadding="1" cellspacing="1" 
        style="font-family: Tahoma; font-size: xx-small; width: 80%">
                         
                    <tr bgcolor="#b0c4de">
                        <th >
                            Sifno</th>
                        <th >
                            Sif Tarih</th>
                        <th >
                            Malzeme</th>
                        <th >
                            Malzeme ack</th>
                        <th >
                            Talep Miktar</th>
                        <th >
                            Birim</th>
                        <th >
                            Temin Miktar</th>                        
                        <th >
                            Temin Birim</th>
                        <th >
                            Maliyet Mekezi</th>                       
                        <th >
                            Acil Mi?</th>
                        <th >
                            Cihaz Kodu</th>
                        <th >
                            Fat No</th>
                        <th >
                            İşlem</th>
                    </tr>  
                    <%While dbread.Read()%>  
                   <FORM id="form1" method="post" action="sifislem.aspx?kapat=<%=dbread("sifdetayid")%>&sifno=<%=dbread("sifnodetay")%>">      
                    <tr bgcolor="#b0c4de">
                        <th class="style1">
                            <%=dbread("sifnodetay")%> </th>
                        <th class="style1">
                            <%=FormatDateTime(dbread("sifdetaytarih"),2)%> </th>
                        <th class="style1" >
                            <%=stokad(dbread("malzemekod"))%> </th>
                        <th class="style1" >
                            <%=dbread("malzemetanim")%></th>
                        <th class="style1" >
                            <%=dbread("talepmiktar")%> </th>                       
                        <th class="style1" >
                            <%=birimad(dbread("birim"))%> </th>
                        <th class="style1" >
                        <input id="Text1" name="teminmiktar" size="25" type="text" value="<%=dbread("talepmiktar")%>" /></th>
                        <th class="style1">                           
                         <SELECT id="select1" name="birim" 
                                style="HEIGHT: 18px; WIDTH: 45px; font-size: xx-small; font-family: Tahoma;" 
                               >
                            <OPTION selected value="<%=dbread("birim")%>"><font size="2" face="tahoma"><%=birimad(dbread("birim"))%></font></OPTION>
                            <%  objConn = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("satinalma.mdb") & ";")
                                objConn.Open()
                                strSQL = "SELECT * from tbl_birim"
                                dbcomm = New OleDbCommand(strSQL, objConn)
                                dbbirim = dbcomm.ExecuteReader()
                                While (dbbirim.Read())%>
                            <OPTION value="<%=dbbirim("code")%>" ><%=dbbirim("name")%></OPTION>>
		                    <%
		                        End While
		                        objConn.Close()
		                        %>
		                 </SELECT>
                        </th>    
                        <th class="style1" >
                            <%=csad(dbread("costcenter"))%> </th>                       
                        <th class="style1" >
                              <%=dbread("acilmi")%> </th>
                        <th class="style1" >
                             <%=mchad(dbread("cihazkod"))%></th>
                        <th class="style1" >
                             <input id="fatno" name="fatno" size="25" type="text" value="<%=dbread("fatno")%>"/></th>
                        <th class="style1" >
                        <input type="submit" name="sifkapat" value="Kapat" 
                              style="height: 17px; font-size: xx-small" />
                           </th>
                    </tr>   
                    </form>
                    <%End While%>         
      
</table>

</body>