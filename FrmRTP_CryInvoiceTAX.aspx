<%@ Page Language="VB" AutoEventWireup="false" CodeFile="FrmRTP_CryInvoiceTAX.aspx.vb" Inherits="ALL_Report_Frm_Acccode_items_report" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <meta http-equiv="content-type" content="text/html; charset=UTF-8">
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable = no">
    <link href="../css/bootstrap.css" rel="stylesheet" type="text/css" />
    <link href="Phetsarath_Font.css" rel="stylesheet" type="text/css" />
   
 
    <style type="text/css">
 


        .style21
        {
            height: 21px;
            font-size: 11pt;
            text-align: left;
        }
        .style23
        {
            height: 21px;
            width: 396px;
            font-size: 11pt;
            text-align: left;
        }
        .style25
        {
            font-size: 11pt;
        }
        .style29
        {
        
            font-size: 11pt;
        }
        .style30
        {
            font-size:11pt;
        }
        .style32
        {
            text-align: left;
        }
        .style35
        {
        	font-size:11pt;
            text-align: left;
            height: 30px;
        }
        .style36
        {
            height: 26px;
        }
        .style37
        {
            text-align: right;
         
        }
        .style38
        {
            width: 50px;
        }

   </style>
</head>
<body onload="window.print()" >
 <img src="../images/A01.jpg" style="position: absolute; top: 0; left: 0; width: 100%; height: 100%; z-index: -1;" />
    <form id="form1" runat="server" >

      <div>
   
       <table style="width:100%; font-size :11pt;">
      <tr>
      <td style="width:4%"></td>
                <td  style="width:20%">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Image ID="Image1" runat="server" Height="64px" 
                        ImageUrl="~/Files/ap22.ico" 
                        Width="79px" />
                    <br />
                    <br />
                    </td >
                <td align="center" style="width:45%">
                    <asp:Label ID="Label43" runat="server" CssClass="style69" 
                        Text="ໃບເກັບເງິນ/Tax Invoice" style="font-size: 22pt; font-weight: bold"></asp:Label>
                    <span class="style32"><span class="style36">
                    <span lang="en-us"><span class="style29">
                    <br />
                    <asp:Label ID="Label59" runat="server" CssClass="style69" 
                        Text="(ອາກອນູນຄ່າເພີ່ມ Vat)" style="font-size: 11pt; "></asp:Label>
                    </span></span></span></span></td>
                
                          <br />
                          <br />
                          <td>
        <table style="width:100%;" border="1" >
                        <tr >
                                 <td class="style29" align=left>
                                     <span lang="en-us">&nbsp;<b><font size="2"><asp:Label ID="Label49" 
                                         runat="server"  Text="ເລກທີ :" Font-Size="11pt"></asp:Label></font></b></span><font 
                                         size="2"><b>&nbsp;<asp:Label ID="lbl_InvoiceNo" runat="server" Font-Size="11pt"></asp:Label><br />
                                     </b>
                                     <span lang="en-us">&nbsp;<b> <asp:Label ID="Label51" runat="server" 
                                         Text="ວັນທີ :" Font-Size="11pt"></asp:Label></b></span><b>&nbsp;<asp:Label 
                                         ID="lbl_InvoiceDate" runat="server" Font-Size="11pt"></asp:Label></font></td>
                        </tr>
                         </table>       
                </td>
                 <td style="width:4%"></td>
            </tr>
            </table>                             
        </div>

     <table style="width: 100%; font-family: 'Saysettha OT'; font-size: 11pt;" >
            <tr style="font-size :11pt;">
                 <td style="width:4%;"></td>
                <td align="center">    
                <table style="width: 100%; height: 650px; font-family: 'Saysettha OT'; font-size: 11pt;" border="1" >
                 <tr style="font-size :11pt;">
         <td class="style23" colspan="2">
            
             &nbsp;<asp:Label ID="Label75" runat="server" Text="ຊື່ວິສາຫະກິດ (ຜູ້ຂາຍ): " 
                 Font-Bold="True" class="style29" ></asp:Label>
                                <span lang="en-us">
             <br />
             &nbsp;<asp:Label ID="lbl_Name" runat="server"  class="style29"  Font-Bold="True"></asp:Label>
                                <br />
             &nbsp;</span><asp:Label ID="Label76" runat="server" class="style29"  Text="ເລກປະຈຳຕົວຜູ້ເສຍອາກອນ :"></asp:Label>
                    <span lang="en-us">
                    &nbsp;<asp:Label ID="lbl_TIN1" runat="server" class="style29" ></asp:Label>
                                <br />
             &nbsp;</span><asp:Label ID="Label77" runat="server" class="style29"  Text="ທີ່ຢູ່ :"></asp:Label>
                    <span lang="en-us"> 
                    &nbsp;<asp:Label ID="lbl_Address1" runat="server"  class="style29" ></asp:Label>
                                <br />
             &nbsp;</span><asp:Label ID="Label78" runat="server" class="style29"  Text="Tel :"></asp:Label>
                    <span lang="en-us">
                    &nbsp;<asp:Label ID="lbl_Tel1" runat="server" class="style29" ></asp:Label>
                                <br />
             &nbsp;</span><asp:Label ID="Label79" runat="server"  class="style29" Text="ຊື່ທະນາຄານ :"></asp:Label>
                    <span lang="en-us">&nbsp;<asp:Label ID="lbl_Bnk_nm" runat="server"  class="style29" ></asp:Label>
                                <br />
             &nbsp;</span><asp:Label ID="Label31" runat="server"  class="style29" Text="ຊື່ບັນຊີ ທ/ຄ :"></asp:Label>
                    <span lang="en-us">
                    &nbsp;<asp:Label ID="lbl_Bnk" runat="server" class="style29" ></asp:Label>
                                <br />
             &nbsp;</span><asp:Label ID="Label80" runat="server"  class="style29" Text="ເລກບັນຊີ ທ/ຄ :"></asp:Label>
                    <span lang="en-us">
                    &nbsp;<asp:Label ID="lbl_Acc" runat="server" class="style29" ></asp:Label>
                    </span>
                     </td>
         <td class="style21" colspan="4">
                    <span lang="en-us"> 
                                &nbsp;<asp:Label ID="Label84" runat="server"  class="style29" 
                 Text="ຊື່ວິສາຫະກິດ (ຜູ້ຊື້):"></asp:Label>
                    &nbsp;<br />
                    &nbsp;<asp:Label ID="lbl_PayerAccountNm" runat="server" class="style29" ></asp:Label>
                      <br />
                    &nbsp;<asp:Label ID="Label85" runat="server" class="style29"  Text="ເລກປະຈຳຕົວຜູ້ເສຍອາກອນ :"></asp:Label>
                    &nbsp;<asp:Label ID="lbl_TIN" runat="server" class="style29" ></asp:Label>
                      <br />
                    &nbsp;<asp:Label ID="Label86" runat="server" class="style29"  Text="ທີ່ຢູ່ :"></asp:Label>
                    &nbsp;<asp:Label ID="lbl_Address" runat="server" class="style29" ></asp:Label>
                      <br />
                    &nbsp;<asp:Label ID="Label81" runat="server" class="style29"  Text="Tel :"></asp:Label>
                    &nbsp;<asp:Label ID="lbl_Tel" runat="server" class="style29" ></asp:Label>
                      <br />
                    &nbsp;<asp:Label ID="Label29" runat="server"  class="style29" Text="ຊື່ບັນຊີ ທ/ຄ :"></asp:Label>
                    &nbsp;<asp:Label ID="lbl_BankAccountNm" runat="server" class="style29" ></asp:Label>
                      <br />
                    &nbsp;<asp:Label ID="Label82" runat="server" class="style29"  Text="ເລກບັນຊີ ທ/ຄ :"></asp:Label>
                    &nbsp;<asp:Label ID="lbl_BankAccountNo" runat="server" class="style29" ></asp:Label>
                      <br />
                    &nbsp;<asp:Label ID="Label83" runat="server" class="style29"  Text="ຮູບແບບການສຳລະສະສາງ :"></asp:Label>
                    &nbsp;<asp:Label ID="lbl_Typepay" runat="server" class="style29" ></asp:Label>
                    &nbsp;</span></td>
         </tr>
            <tr style="font-size :10pt;  font-family: 'Saysettha OT';">
                <td align="center"  style="width:5%; height: 10px">
                    <asp:Label ID="Label6" runat="server" Text="ລ/ດ" CssClass="style25"></asp:Label>
                    </td>

                 <td align="center" style="width: 45%; height: 10px">
                    <asp:Label ID="Label61" runat="server" Text="ເນື້ອໃນລາຍການ" CssClass="style25"></asp:Label>
                </td>
                <td align="center" style="width: 8%; height: 10px">
                    <asp:Label ID="Label10" runat="server" Text="ຫົວໜ່ວຍ" CssClass="style25"></asp:Label>
    
                </td>
                     <td align="center" style="width: 8%; height: 10px">
                    <asp:Label ID="Label2" runat="server" Text="ຈຳນວນ" CssClass="style25"></asp:Label>
                </td>
                <td align="center" style="width: 16%; height: 10px">
                    <asp:Label ID="Label3" runat="server" Text="ລາຄາ" CssClass="style25"></asp:Label>
                
                    </td>
                     <td align="center" style="width: 18%; height: 10px">
                    <asp:Label ID="Label17" runat="server" Text="ມູນຄ່າ" CssClass="style25"></asp:Label>
                        
                    </td>
            </tr>                
        <asp:Repeater ID="Repeater1" runat="server">
    <HeaderTemplate>
        
          </HeaderTemplate>
   <ItemTemplate>
            <tr style="height: 100%; ">          
                <td align="center"  class="style29" style="width: 60px"   >
                           <asp:Label ID="lblRowNumber" Text='<%# Container.ItemIndex + 1 %>' runat="server" Font-Size="10pt" />                               
                </td>
                <td align="left" class="style29" >
                    
                    <asp:Literal ID="literal1" runat="server" Text='<%#Eval("ItemNm")%>' ></asp:Literal>
                    
                </td>
                     <td align="center" class="style29" style="width: 60px"  >      
                    <asp:Label ID="Label5" runat="server"  ></asp:Label>
                         &nbsp;
                 <asp:Literal ID="Literal6" runat="server" 
                        Text='<%#Eval("Unit", "{0:#,##0.00}")%>'></asp:Literal> 
                   </td>
                <td align="center"  style="width: 60px" class="style29" >                  
                    <asp:Literal ID="literal2" runat="server" Text='<%#Eval("Qty")%>'></asp:Literal>
                </td>
           
                   <td align="right"  style="width: 60px" class="style29" >      
                    <asp:Label ID="Label4" runat="server"  ></asp:Label>
                         &nbsp;
                 <asp:Literal ID="Literal3" runat="server" 
                        Text='<%#Eval("Price", "{0:#,##0.00}")%>'></asp:Literal> 
                   </td>
                   <td align="right"  style="width: 70px" class="style29" >      
                    <asp:Label ID="Label16" runat="server"  ></asp:Label>
                         &nbsp;
                 <asp:Literal ID="Literal7" runat="server" 
                        Text='<%#Eval("Amonth", "{0:#,##0.00}")%>'></asp:Literal> 
                   </td>
                  
            </tr>
            
           </ItemTemplate>
    <FooterTemplate>
     
       
        </FooterTemplate>
</asp:Repeater>

 <tr align="right" >
     <td colspan="2" align="right" rowspan="4" > ​<span class="style29">​​​​&nbsp;</span><span lang="en-us">&nbsp;</span></td>
     </tr>
     <tr>
     <td align="right" colspan="3"  class="style29"  >
         <asp:Label ID="Label11" runat="server" 
             Text="ລວມມູນຄ່າບໍ່ມີອາກອນ" Font-Size="10pt"></asp:Label></td>
     <td align=right> <asp:Label ID="lbl_AmountNet" runat="server" Text="0" Font-Size="10pt"></asp:Label> </td>
     </tr>
     <tr>
     <td colspan="3" align="right" style="height:10px;"> ​<asp:Label ID="Label25" runat="server" 
             Text="ອັດຕາອາກອນມູນຄ່າເພີ່ມ" Font-Size="10pt"></asp:Label> 
             &nbsp; 
             <asp:Label ID="lbl_Tax_Rate" runat="server" 
             Text="0" Font-Size="10pt"></asp:Label> <span lang="en-us">&nbsp;<asp:Label ID="Label9" 
             runat="server" Text="%" Font-Size="10pt"></asp:Label></span>
             </td>
     <td align=right> <asp:Label ID="lbl_AmountTax" runat="server" Text="0" Font-Size="10pt"></asp:Label> </td>
     </tr>
     <tr>
             
     <td colspan="3" align="right" style="height:10px;"> ​<asp:Label ID="Label20" runat="server" 
             Text="ລວມມູນຄ່າທັງໝົດ (LAK) " Font-Size="10pt" ></asp:Label></td>
     <td class="style37" colspan="1">  
         <asp:Label ID="lbl_AmountTotal" runat="server" 
             Text="0" Font-Size="10pt" ></asp:Label> </td>
     </tr>
         <tr>
     <td colspan="6" class="style35" > 
     <span lang="en-us"><asp:Label ID="lbl_AmtLAk0" runat="server" Font-Italic="True" Font-Size="10pt">ຈຳນວນຂຽນເປັນຕົວໜັງສື 
         :</asp:Label>&nbsp;
     <asp:Label ID="lbl_AmtLAk" runat="server" Font-Size="10pt"></asp:Label></span><asp:TextBox 
             ID="TextBox1" runat="server" Visible="False" Width="16px" Font-Size="10pt"></asp:TextBox></td>
             
</tr>
</table>

  <table  class="style32" style="width: 100%;"
                           border="1">
                             <tr>
     <td style="width: 50%; height: 120px;" align="center"> 
                                
                    <asp:Label ID="S6" runat="server" 
                        Font-Bold="True" Font-Size="10pt">ຜູ້ຊື້</asp:Label>
                    
                                <br />
                                
                    <asp:Label ID="S7" runat="server" 
                        Font-Bold="False" Font-Size="10pt">(ລາຍເຊັນ ແລະ ຈ້ຳກາ)</asp:Label>
                    
                    <br />
                    <br />
                    
                    <br />
                    
                    <br />
                        <font size="1">
                    <asp:Label ID="Label90" runat="server" 
                            Text="ຊື່ແຈ້ງ:......................................................." 
                            CssClass="style29"></asp:Label>
                                     <br />
                    <asp:Label ID="Label92" runat="server" 
                            Text="ວັນທີ:........................................................." 
                            CssClass="style29"></asp:Label>
                    </font>
                    <br />
                    
         </td>
        <td colspan="2" style="width: 50%" align="center"> 
           <asp:Label ID="Label13" runat="server" 
                        Font-Bold="True" Font-Size="10pt">ຜູ້ຂາຍ</asp:Label>
                    
                                <br />
                                
                    <asp:Label ID="Label15" runat="server" 
                        Font-Bold="False" Font-Size="10pt">(ລາຍເຊັນ ແລະ ຈ້ຳກາ)</asp:Label>
            <br />
            <br />
            <br />
            <br />
                        <font size="1">
                    <asp:Label ID="Label91" runat="server" 
                            Text="ຊື່ແຈ້ງ:......................................................." 
                            CssClass="style29"></asp:Label>
                                     <br />
                    <asp:Label ID="Label93" runat="server" 
                            Text="ວັນທີ:........................................................." 
                            CssClass="style29"></asp:Label>
                    </font>
            <br />
         </td>     
     </tr>
      <tr >
     <td style="border-style: solid none solid solid; border-width: thin; border-color: #000000;"> 
                                
                    <asp:Label ID="Label87" runat="server" 
                        Font-Bold="False" Font-Size="8pt">ໃບອະນຸຍາດນຳໃຊ້ໃບເກັບເງິນສະເພາະ </asp:Label>
                    
                                <br />
                                
                    <asp:Label ID="Label23" runat="server" 
                        Font-Bold="False" Font-Size="8pt">ເລກທີ : </asp:Label>
                    
                    <asp:Label ID="lbl_No" runat="server" 
                        Font-Bold="False" Font-Size="8pt"></asp:Label>
                    ,&nbsp;
                    <asp:Label ID="lbl_SN1" runat="server" 
                        Font-Bold="False" Font-Size="8pt">ລົງວັນທີ :</asp:Label>
                    &nbsp;<asp:Label ID="lbl_Date" runat="server" 
                        Font-Bold="False" Font-Size="8pt"></asp:Label>
                    
                    <br />
                    
         </td>
           
     <td style="border-width: thin; border-color: #000000; border-style: solid none solid none;" 
              class="style38" align="right"> 
                                
                    <asp:Image ID="Image2" runat="server" ImageUrl="~/images/apis_Small.JPG" />
                    
         </td>
           
     <td style="border-style: solid solid solid none; border-width: thin; border-color: #000000; font-family: 'Saysettha OT'; font-size: 8pt; width: 70px;" 
              align="center">  ພັດທະນາໂດຍ ບໍລິສັດເອພີໄອເອັສ ຈໍາກັດ</td>
           
     </tr>
       </table>
   </td>
   <td style="width:4%"></td>
   </tr>
   </table>
    </form>
</body>
</html>
