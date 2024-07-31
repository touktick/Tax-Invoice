Imports System.Data.SqlClient
Imports System.Data
Imports System
Imports System.Diagnostics
Imports System.ComponentModel
Imports System.Configuration
Partial Class ALL_Report_Frm_Acccode_items_report
    Inherits System.Web.UI.Page
    Dim cnstr As String = ""
    Dim cn As New SqlConnection
    Dim da As New SqlDataAdapter
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("Report") = "" Then
            Response.Write("<script language='javascript'>self.close();</script>")
        End If
        If Not IsPostBack Then
 
            Call BindList()
            Load_Office()
            calcTotal()
            TextBox1.Text = lbl_AmountTotal.Text

        End If
        Amount_Later()

    End Sub
  
    Private Sub Load_Office()
        Try
            Dim sql As String = ""
            Dim Supp As String = ""
            Dim da As New SqlDataAdapter
            Dim ds As New DataSet
            With cn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = Session("cnstr")
                .Open()
            End With
            'sql = "SELECT CM.*, IV.Lock, IV.InvoiceNo, IV.InvoiceDate, IV.DueDate, IV.JobNo, IV.PayerAccountID, IV.ItemCount, IV.AmountInvoice, IV.AmountTax, 0, IV.AmountNet, IV.AmountDiscount, IV.PercentDiscount, IV.AmountTotal, IV.AmountPaid, IV.Currency,  " & _
            '            " IV.InvoiceDescription, IV.IsFinished, IV.THB_LAK, IV.USD_LAK, IV.AcRevenue, IV.AcTax, IV.Paid_Date, IV.AcAdv, IV.credit_term, IV.AmountLAK, IV.AcDebit, CM.PayerAccountNm, IV.TIN, IV.Tax_Rate, dbo.AP_Office.*  " & _
            '            " from AP_TAX_InvList IV Inner join AP_Customers CM ON IV.PayerAccountID=CM.CustomerID CROSS JOIN dbo.AP_Office " & _
            '     " where IV.InvoiceNo='" & Trim(Session("ReportID")) & "' and dbo.AP_Office.off_id=N'" & Trim(Session("Off_Id")) & "' "
            'da = New SqlDataAdapter(sql, cn)
            'da.Fill(ds, "AP_A2")
            'If ds.Tables("AP_A2").Rows.Count > 0 Then
            '    Label19.Text = (ds.Tables("AP_A2").Rows(0).Item("PayerAccountNm").ToString)
            '    lbl_TIN.Text = (ds.Tables("AP_A2").Rows(0).Item("TIN").ToString)
            '    lbl_TelNo.Text = (ds.Tables("AP_A2").Rows(0).Item("TelNo").ToString)
            '    lbl_FaxNo.Text = (ds.Tables("AP_A2").Rows(0).Item("FaxNo").ToString)
            '    lbl_Address.Text = (ds.Tables("AP_A2").Rows(0).Item("Address").ToString)

            '    lbl_SN.Text = (ds.Tables("AP_A2").Rows(0).Item("txtSN").ToString)
            '    lbl_NO.Text = (ds.Tables("AP_A2").Rows(0).Item("txtNO").ToString)
            '    lbl_InvoiceNo.Text = (ds.Tables("AP_A2").Rows(0).Item("InvoiceNo").ToString)
            '    lbl_InvoiceDate.Text = Format(CDate(ds.Tables("AP_A2").Rows(0).Item("InvoiceDate")), "dd/MM/yyyy")
            '    lbl_AmountNet.Text = Format(CDbl(ds.Tables("AP_A2").Rows(0).Item("AmountNet")), "#,##0.00")
            '    lbl_AmountTax.Text = Format(CDbl(ds.Tables("AP_A2").Rows(0).Item("AmountTax")), "#,##0.00")
            '    lbl_AmountTotal.Text = Format(CDbl(ds.Tables("AP_A2").Rows(0).Item("AmountTotal")), "#,##0.00")
            '    lbl_Tax_Rate.Text = (ds.Tables("AP_A2").Rows(0).Item("Tax_Rate").ToString)
            '    '  totalDisc.Text = Format((ds.Tables("AP_A2").Rows(0).Item("AmountInvoice").ToString), "#,##0.00")


            'End If


            sql = ""
            sql = " select * from AP_Office where off_id='" & Trim(Session("Off_Id")) & "' "
            da = New SqlDataAdapter(sql, cn)
            da.Fill(ds, "AP_A2")
            If ds.Tables("AP_A2").Rows.Count > 0 Then
                lbl_Name.Text = (ds.Tables("AP_A2").Rows(0).Item("off_nm").ToString)
                lbl_TIN1.Text = (ds.Tables("AP_A2").Rows(0).Item("ACCNOTAX").ToString)
                lbl_No.Text = (ds.Tables("AP_A2").Rows(0).Item("txtNO").ToString)
                lbl_Date.Text = (ds.Tables("AP_A2").Rows(0).Item("txtDT").ToString)
                '  lbl_SN.Text = (ds.Tables("AP_A2").Rows(0).Item("txtSN").ToString)
                lbl_Address1.Text = (ds.Tables("AP_A2").Rows(0).Item("off_StrtL").ToString)
                lbl_Tel1.Text = (ds.Tables("AP_A2").Rows(0).Item("Tel").ToString)
                lbl_Bnk_nm.Text = (ds.Tables("AP_A2").Rows(0).Item("Bnk_nm").ToString)
                lbl_Bnk.Text = (ds.Tables("AP_A2").Rows(0).Item("Bnk1").ToString)
                lbl_Acc.Text = (ds.Tables("AP_A2").Rows(0).Item("ACC1").ToString)
            End If
            sql = ""
            sql = " SELECT IV.InvoiceNo, IV.InvoiceDate, PA.PayerAccountNm, PA.Address, PA.TelNo, PA.FaxNo, PA.ContractName, PA.BankNm, PA.BankAddress, PA.BankAccountNm, PA.BankAccountNo, PA.SwiftCode, PA.BankTel, IV.Type_pay, " & _
                  " IV.TAX_Rate,  IV.AmountInvoice, IV.AmountTax, IV.AmountTotal, IV.AmountNet, PA.TIN, IV.Letter_Amt  FROM     dbo.AP_TAX_InvList AS IV LEFT OUTER JOIN   dbo.AP_Customers AS PA ON PA.CustomerID = IV.PayerAccountID where IV.InvoiceNo = '" & Session("Biil_No") & "'   "
            da = New SqlDataAdapter(sql, cn)
            da.Fill(ds, "AP")
            If ds.Tables("AP").Rows.Count > 0 Then
                lbl_InvoiceNo.Text = (ds.Tables("AP").Rows(0).Item("InvoiceNo").ToString)
                lbl_InvoiceDate.Text = Format(CDate(ds.Tables("AP").Rows(0).Item("InvoiceDate")), "dd/MM/yyyy")

                lbl_PayerAccountNm.Text = (ds.Tables("AP").Rows(0).Item("PayerAccountNm").ToString)
                lbl_TIN.Text = (ds.Tables("AP").Rows(0).Item("TIN").ToString)
                lbl_Address.Text = (ds.Tables("AP").Rows(0).Item("Address").ToString)
                lbl_Tel.Text = (ds.Tables("AP").Rows(0).Item("TelNo").ToString)
                lbl_BankAccountNm.Text = (ds.Tables("AP").Rows(0).Item("BankAccountNm").ToString)
                lbl_BankAccountNo.Text = (ds.Tables("AP").Rows(0).Item("BankAccountNo").ToString)
                lbl_Typepay.Text = (ds.Tables("AP").Rows(0).Item("Type_pay").ToString)
                lbl_AmtLAk.Text = (ds.Tables("AP").Rows(0).Item("Letter_Amt").ToString)
                lbl_AmountNet.Text = Format(CDbl(ds.Tables("AP").Rows(0).Item("AmountNet")), "#,##0.00")
                lbl_AmountTax.Text = Format(CDbl(ds.Tables("AP").Rows(0).Item("AmountTax")), "#,##0.00")
                lbl_AmountTotal.Text = Format(CDbl(ds.Tables("AP").Rows(0).Item("AmountTotal")), "#,##0.00")
                lbl_Tax_Rate.Text = (ds.Tables("AP").Rows(0).Item("Tax_Rate").ToString)
            End If


        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub
    Protected Sub calcTotal()
        'Dim sum As Double = 0
        'Dim i As Integer

        'sum = 0
        'For i = 0 To Repeater1.Items.Count - 1
        '    If (Repeater1.Items(i).ItemType = ListItemType.Item) Or (Repeater1.Items(i).ItemType = ListItemType.AlternatingItem) Then
        '        sum += CType(Repeater1.Items(i).FindControl("literal2"), Literal).Text
        '    End If
        'Next
        'lbl_TotalQty.Text = Format(CDbl(sum), "#,###")

        'sum = 0
        'For i = 0 To Repeater1.Items.Count - 1
        '    If (Repeater1.Items(i).ItemType = ListItemType.Item) Or (Repeater1.Items(i).ItemType = ListItemType.AlternatingItem) Then
        '        sum += CType(Repeater1.Items(i).FindControl("Literal4"), Literal).Text
        '    End If
        'Next
        'Bill_Amt.Text = Format(CDbl(sum), "#,##0.00")

        
    End Sub
    Private Sub BindList()
        Try
            Dim constr As String = Session("cnstr")
            Using con As New SqlConnection(constr)
                Using cmd As New SqlCommand()
                    cmd.CommandText = Session("Report")
                    cmd.Connection = con
                    Using sda As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        sda.Fill(dt)
                        Repeater1.DataSource = dt
                        Repeater1.DataBind()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

    Protected Sub Repeater1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.RepeaterCommandEventArgs) Handles Repeater1.ItemCommand




    End Sub
    Dim curr, CRR As String
    Private Sub Amount_Later()
        curr = ""
        'If Label50.Text = "THB" Then
        '    lbl_AmtLAk.Text = Letter_amt(TextBox1) & " ບາດ"

        'ElseIf Label50.Text = "LAK" Then
        '    lbl_AmtLAk.Text = Letter_amt(TextBox1) & " ກີບ"
        'Else
        '    lbl_AmtLAk.Text = Letter_amt(TextBox1) & " ໂດລາ"
        'End If
        lbl_AmtLAk.Text = Letter_amt(TextBox1) & " ກີບ"
    End Sub

    Public Function Piase(ByVal pi As String) As String
        Dim Istr As String
        Istr = ""
        Dim nos As Object
        nos = pi
        Select Case nos
            Case 0 To 9 : Istr = Choose(nos + 1, "ສູນ", "ໜຶ່ງ", "ສອງ", "ສາມ", "ສີ່", "ຫ້າ", "ຫົກ", "ເຈັດ", "ແປດ", "ເກົ້າ")
        End Select
        Piase = Istr
    End Function

    Public Function Letter_amt(ByVal Txt As TextBox, Optional ByVal CurrKIP As Boolean = False) As String
        If Val(TextBox1.Text) <> 0 Then
            Letter_amt = CMoney(Format(CDbl(Txt.Text), "##0.00"))
        Else
            Letter_amt = ""
        End If
    End Function
    Public Function Inwords(ByVal str1 As String) As String
        Dim Istr As String
        Istr = ""
        Dim nos As Object
        nos = CDbl(str1)

        Select Case nos
            Case 0 To 19
                Istr = Choose(nos + 1, "", "ໜຶ່ງ", "ສອງ", "ສາມ´", "ສີ່", "ຫ້າ", "ຫົກ", "ເຈັດ", "ແປດ", "ເກົ້າ", "ສິບ", "ສິບເອັດ", "ສິບສອງ", "ສິບສາມ´", "ສິບສີ່", "ສິບຫ້າ", "ສິບຫົກ", "ສິບເຈັດ", "ສິບແປດ", "ສິບເກົ້າ")
            Case 20, 30, 40, 50, 60, 70, 80, 90
                Istr = Choose(nos / 10, "ກີບ", "ຊາວ", "ສາມສິບ", "ສີ່ສິບ", "ຫ້າສິບ", "ຫົກສິບ", "ເຈັດສິບ", "ແປດສິບ", "ເກົ້າສິບ") 'RAVI
            Case 21, 31, 41, 51, 61, 71, 81, 91
                Istr = Choose(nos / 10, "ກີບ", "ຊາວເອັດ", "ສາມສິບເອັດ", "ສີ່ສິບເອັດ", "ຫ້າສິບເອັດ", "ຫົກສິບເອັດ", "ເຈັດສິລເອັດ", "ແປດສິບເອັດ", "ເກົ້າສິບເອັດ") 'RAVI
            Case 22 To 30, 32 To 40, 42 To 50, 52 To 60, 62 To 70, 72 To 80, 82 To 99
                Istr = Inwords(Left(nos, 1) & 0) & Inwords(Right(nos, 1))
            Case 101 To 999
                If CInt(nos) <> 0 Then
                    Istr = Inwords(Left(nos, 1)) & "ຮ້ອຍ" & Inwords(Right(nos, 2))
                Else
                    Istr = Inwords(Left(nos, 1)) & "ຮ້ອຍ" & Inwords(Right(nos, 2))
                End If
            Case 1001 To 9999
                Istr = Inwords(Left(nos, 1)) & "ພັນ" & Inwords(Right(nos, 3))
            Case 10000 To 99999
                Istr = Inwords(Mid(nos, 1, 2)) & "ພັນ" & Inwords(Right(nos, 3))
            Case 100000 To 999999
                Istr = Inwords(Left(nos, 1)) & "ແສນ" & Inwords(Right(nos, 5))
            Case 1000000 To 9999999
                Istr = Inwords(Left(nos, 1)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 10000000 To 99999999
                Istr = Inwords(Left(nos, 2)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 100000000 To 999999999
                Istr = Inwords(Left(nos, 3)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 1000000000 To 9999999999.0#
                Istr = Inwords(Left(nos, 4)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 10000000000.0# To 99999999999.0#
                Istr = Inwords(Left(nos, 5)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 100000000000.0# To 999999999999.0#
                Istr = Inwords(Left(nos, 6)) & "ລ້ານ" & Inwords(Right(nos, 6))
            Case 100
                Istr = Inwords(Left(nos, 1)) & "ຮ້ອຍ"
            Case 1000
                Istr = Inwords(Left(nos, 1)) & "ພັນ"
            Case 100000
                Istr = Inwords(Left(nos, 1)) & "ແສນ"
            Case 10000000
                Istr = Inwords(Left(nos, 2)) & "ລ້ານ"
        End Select
        Inwords = Istr
    End Function

    Public Function CMoney(ByVal strr As String) As String
        Dim rs As Object
        Dim i As Integer
        Dim A As String
        A = ""
        rs = (Right(strr, (Len(strr) - InStr(1, strr, "."))))
        'If Val(strr) = 0 Then
        'Exit Function
        'End If
        If InStr(1, strr, ".") <> 0 Then
            If Len(Right(strr, (Len(strr) - InStr(1, strr, ".")))) = 1 Then
                If Val(Mid(strr, 1, InStr(1, strr, "."))) = 0 Then
                    CMoney = Piase(rs * 10) & curr
                Else
                    CMoney = Inwords(Left(strr, InStr(1, strr, ".") - 1)) & " ຈຸດ" & Piase(rs * 10) & curr
                End If
            Else
                For i = 1 To Len(Right(strr, (Len(strr) - InStr(1, strr, "."))))
                    A = A & Piase(Mid(Right(strr, (Len(strr) - InStr(1, strr, "."))), i, 1))
                Next
                If Val(Right(strr, (Len(strr) - InStr(1, strr, ".")))) > 0 Then
                    If Val(Left(strr, InStr(1, strr, ".") - 1)) = 0 Then
                        CMoney = "ສູນ" & " ຈຸດ" & A & curr
                    Else
                        CMoney = Inwords(Left(strr, InStr(1, strr, ".") - 1)) & " ຈຸດ" & A & curr
                    End If
                Else
                    CMoney = Inwords(Left(strr, InStr(1, strr, ".") - 1)) & curr
                End If
                '       End If
            End If
        Else
            CMoney = Inwords(strr)
        End If
    End Function

End Class
