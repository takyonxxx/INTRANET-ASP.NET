<head>
   
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
    <style type="text/css">
        #Button2
        {
            font-family: Tahoma;
            font-size: x-small;
            font-weight: 700;
            height: 21px;
        }
        .style30
        {
            text-decoration: none;
        }
    </style>
</head>
<table   align="center" width="60%">
    <tr bgcolor="Gainsboro">
        <td style="font-family: Tahoma; font-size: x-small; font-weight: 700; color: #FFFFFF; background-color: #003366" >
            &nbsp;
            <A HREF="main.aspx" style="color: #FFFFFF" class="style30">Main Page</A></td>
        <td style="font-family: Tahoma; font-size: x-small; font-weight: 700; color: #FFFFFF; background-color: #003366" >
             <input id="Text1" type="text" value="<%=Session("sifno")%>" size="25"/>
             <input id="Button2" type="button" value="FIND PO" onclick="setLocation('Text1');"  /></td>
        <td style="font-family: Tahoma; font-size: x-small; font-weight: 700; color: #FFFFFF; background-color: #003366" >
            &nbsp;
            <A HREF="sifrapor.aspx" style="color: #FFFFFF" class="style30">Search PO</A>&nbsp; </td>
    </tr>   
</table>
