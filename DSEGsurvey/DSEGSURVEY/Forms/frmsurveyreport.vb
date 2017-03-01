Imports System.ComponentModel
Imports System.Threading
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Public Class frmsurveyreport
    Dim ripindex As Integer = 0
    Dim selyear As Integer = 0
    Dim ryear As Integer = 0
    Dim gtot As Integer = 0
    Dim ccol As Integer = 0
    Dim tr1 As System.Threading.Thread
    Private Delegate Function ReturnDelegate() As Object
    Private row As New DataGridViewRow
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    Private Sub cbosy_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbosy.SelectedIndexChanged
        If cbosy.Text <> Nothing And cboryear.Text <> Nothing Then
            cbosy.Enabled = False
            cboryear.Enabled = False
            cbocol.Enabled = False
            btnexport.Enabled = False


            setter()
        End If
    End Sub

    Private Sub cboryear_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboryear.SelectedIndexChanged
        If cbosy.Text <> Nothing Then
            cbosy.Enabled = False
            cboryear.Enabled = False
            cbocol.Enabled = False
            btnexport.Enabled = False

            setter()
        End If
    End Sub
    Sub addheader()
        Dim col As New DataGridViewTextBoxColumn
        col.DataPropertyName = "ITEMS"
        col.HeaderText = "ITEMS"
        col.Name = "ITEMS"
        col.Width = 350
        col.SortMode = DataGridViewColumnSortMode.NotSortable
        col.ReadOnly = True
        dg1.Columns.Add(col)

        Dim col2 As New DataGridViewTextBoxColumn
        col2.DataPropertyName = "DETAILS"
        col2.HeaderText = "DETAILS"
        col2.Name = "DETAILS"
        col2.Width = 200
        col2.ReadOnly = True
        col2.SortMode = DataGridViewColumnSortMode.NotSortable
        dg1.Columns.Add(col2)

        Dim col3 As New DataGridViewTextBoxColumn
        col3.DataPropertyName = "DETAILS1"
        col3.HeaderText = " "
        col3.Name = "DETAILS1"
        col3.ReadOnly = True
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.Width = 200
        dg1.Columns.Add(col3)
    End Sub
    Sub setter()

        dg1.Rows.Clear()
        dg1.Columns.Clear()
        ripindex = cbosy.SelectedIndex

        addheader()
        selyear = Strings.Left(cbosy.Text, 4)
        ryear = cboryear.Text

        Dim gt As New ADODB.Recordset
        sopen(gt)
        If cbocol.Text <> Nothing Then
            gt.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            gt.Open("select count(*) as cnt from respondent where myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & "", db)
        End If

        If gt.EOF = False Then
            gtot = gt.Fields("cnt").Value
            lblgtot.Text = gtot
        End If
        gt.Close()


        If cbosy.Text <> Nothing Then
            lblbarep.Visible = True
            pbar1.Visible = True
            ' btnexport.Visible = False
            tr1 = Nothing
            If tr1 Is Nothing Then
                tr1 = New Thread(Sub() getage())
                tr1.IsBackground = True
                tr1.Start()
            End If

            'getage()
            ' getgender()
            'getreport()
        End If

    End Sub

    Private Sub frmsurveyreport_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        tr1.Abort()
    End Sub

    Private Sub frmsurveyreport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        dg1.Rows.Clear()
        cbosy.Items.Clear()
        cbocol.Items.Clear()

        Dim x As Integer = 0
        x = getdatetime(3)
        For j = 2000 To x
            cbosy.Items.Add(j & "-" & j + 1)
        Next
        cbocol.Items.Add("")
        Dim gcol As New ADODB.Recordset
        sopen(gcol)
        gcol.Open("Select * from college order by GroupName", db)
        While Not gcol.EOF
            cbocol.Items.Add(gcol.Fields("GroupName").Value)
            gcol.MoveNext()
        End While
        gcol.Close()
        'cboryear.SelectedIndex = 0
        ' cbosy.SelectedIndex = 0
    End Sub
    Private Overloads Function AddRow() As Integer

        If InvokeRequired Then
            Return CInt(Invoke(New ReturnDelegate(AddressOf AddRow)))
        Else
            Return dg1.Rows.Add(row)
        End If

    End Function
    Private Sub SetLabelText(ByVal text As String)
        ' Label1.BeginInvoke(Sub() Me.Label1.Text = text)
    End Sub
    Sub getage()
        Dim atot As Integer = 0
        Dim pertot As Double = 0

        Dim per As Double = 0


        Dim grm As New ADODB.Recordset
        sopen(grm)


        pbar1.Value += 4
        lblbarep.Text = "Please wait, Getting Frequency/Percentage for Age ..."
        'Dim row As DataGridViewRow
        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "School Year"
        row.Cells(1).Value = cbosy.Text
        row.Cells(2).Value = Nothing
        row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        AddRow()


        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "Age"
        row.Cells(1).Value = "Frequency"
        row.Cells(2).Value = "Percentage"
        row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
        Addrow()

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "16-17"
        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=16 or age=17) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where (age=16 or age=17) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & " ", db)
        End If

        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If

        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        AddRow()
        grm.Close()

        pbar1.Value += 6
        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "18-19"

        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=18 or age=19) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where (age=18 or age=19) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & " ", db)
        End If
        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If
        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        AddRow()
        grm.Close()




        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "20-21"

        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=20 or age=21) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where (age=20 or age=21) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & " ", db)
        End If
        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If
        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        AddRow()
        grm.Close()

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "22-23"

        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=22 or age=23) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where (age=22 or age=23) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & " ", db)
        End If
        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If
        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        Addrow()
        grm.Close()

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "24 & above"
        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and age>=24 and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where age>=24 and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & " ", db)
        End If
        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If
        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        Addrow()
        grm.Close()

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "TOTAL"
        row.Cells(1).Value = atot
        row.Cells(2).Value = Math.Round(pertot)

        Addrow()

        pbar1.Value += 5
        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Style.BackColor = Color.Gray
        row.Cells(1).Style.BackColor = Color.Gray
        row.Cells(2).Style.BackColor = Color.Gray

        Addrow()
        tr1 = Nothing
        If tr1 Is Nothing Then
            tr1 = New Thread(Sub() getgender())
            tr1.IsBackground = True
            tr1.Start()
        End If
    End Sub
    Sub getgender()
        Dim atot As Integer = 0
        Dim pertot As Double = 0

        Dim per As Double = 0


        Dim grm As New ADODB.Recordset
        sopen(grm)

        pbar1.Value += 5
        lblbarep.Text = "Please wait, Getting Frequency/Percentage for Gender ..."
        'Dim row As DataGridViewRow

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "Gender"
        row.Cells(1).Value = "Frequency"
        row.Cells(2).Value = "Percentage"
        row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
        Addrow()

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "Male"
        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and gender='Male' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where gender='Male' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & " ", db)
        End If

        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If

        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        Addrow()
        grm.Close()

        pbar1.Value += 5
        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "Female"

        If cbocol.Text <> Nothing Then
            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and gender='Female' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=" & svid & "", db)
        Else
            grm.Open("select count(*) as cnt from respondent where gender='Female' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=" & svid & "", db)
        End If
        If gtot > 0 Then
            If grm.EOF = False Then
                row.Cells(1).Value = grm.Fields("cnt").Value
                per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                row.Cells(2).Value = Math.Round(per, 2)
            End If
        Else
            row.Cells(1).Value = grm.Fields("cnt").Value
            row.Cells(2).Value = Nothing
        End If
        pertot = pertot + per
        atot = atot + grm.Fields("cnt").Value
        Addrow()
        grm.Close()


        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "TOTAL"
        row.Cells(1).Value = atot
        row.Cells(2).Value = Math.Round(pertot)

        Addrow()

        pbar1.Value += 5
        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Style.BackColor = Color.Gray
        row.Cells(1).Style.BackColor = Color.Gray
        row.Cells(2).Style.BackColor = Color.Gray

        Addrow()
        tr1 = Nothing
        If tr1 Is Nothing Then
            tr1 = New Thread(Sub() getreport())
            tr1.IsBackground = True
            tr1.Start()
        End If
    End Sub
    Sub getreport()
        Dim per As Double = 0
        Dim averagewmean As Double = 0

        Dim grm As New ADODB.Recordset
        sopen(grm)

        'Dim row As DataGridViewRow
        pbar1.Value += 5

        row = New DataGridViewRow()
        row.CreateCells(dg1)
        row.Cells(0).Value = "Items"
        row.Cells(1).Value = "Weighted Mean"
        row.Cells(2).Value = "Interpretation"
        row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
        Addrow()


        pbar1.Value += 5

        Dim gcat As New ADODB.Recordset
        Dim gk As New ADODB.Recordset
        sopen(gk)
        sopen(gcat)

        Dim mcnt As Integer = 0
        gcat.Open("Select COUNT(*) as cnt from tab_surveycategory where surveyid=" & svid & "", db)
        If gcat.EOF = False Then

            mcnt = gcat.Fields("cnt").Value
        End If
        gcat.Close()

        Dim myadd As Integer = 0
        myadd = 55 / mcnt


        gcat.Open("Select * from tab_surveycategory where surveyid=" & svid & "", db)
        If gcat.EOF = False Then

            While Not gcat.EOF
                lblbarep.Text = "Please wait... Computing Weighted Mean for ITEM - " & gcat.Fields("partname").Value & " ..."

                Dim titem As Integer = 0
                Dim titlimit As Integer = 0
                Dim myper As Double = 0
                Dim mcid As String = Nothing


                gk.Open("Select * from tab_surveykey where surveyid=" & svid & " order by keyname desc ", db)
                If gk.EOF = False Then
                    titlimit = gk.Fields("keyname").Value
                End If
                gk.Close()


                mcid = gcat.Fields("catid").Value
                titem = gcat.Fields("itemcount").Value
                myper = gcat.Fields("percentage").Value



                Dim totitem() As Double
                Dim x As Integer
                ReDim Preserve totitem(titem)
                x = 0
                For i = 1 To titem


                    Dim itemmean As Double = 0
                    For j = 1 To titlimit
                        Dim tmpvar As Double = 0
                        If cbocol.Text <> Nothing Then
                            grm.Open("select count(distinct(r.rno)) as cnt from respondent r, tab_surveyans rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.qid=" & i & " and catid='" & mcid & "'  and rp1.rate=" & j & " and r.rno=rp1.rno and r.myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and r.surveyid=" & svid & "", db)
                        Else
                            grm.Open("select count(*) as cnt from respondent r, tab_surveyans rp1 where rp1.qid=" & i & " and catid=" & mcid & " and rp1.rate=" & j & " and r.rno=rp1.rno and r.myear=" & selyear & " and `year`=" & ryear & " and r.surveyid=" & svid & " ", db)
                        End If

                        If grm.EOF = False Then
                            tmpvar = tmpvar + (grm.Fields("cnt").Value * j)
                            ' row.Cells(1).Value = grm.Fields("cnt").Value
                            ' per = Math.Round((grm.Fields("cnt").Value / gtot) * 100, 2)
                            ' row.Cells(2).Value = Math.Round(per, 2)
                        End If

                        itemmean = itemmean + tmpvar

                        grm.Close()
                    Next
                    row = New DataGridViewRow()
                    row.CreateCells(dg1)
                    row.Cells(0).Value = i

                    If gtot > 0 Then
                        totitem(x) = Math.Round(itemmean / gtot, 2)
                        row.Cells(1).Value = Math.Round(itemmean / gtot, 2)
                    Else
                        row.Cells(1).Value = 0
                        totitem(x) = 0
                    End If
                    Addrow()


                    x += 1
                Next

                Dim totalitemmean As Double = 0

                For x = 0 To UBound(totitem) - 1
                    totalitemmean += totitem(x)
                Next

                row = New DataGridViewRow()
                row.CreateCells(dg1)
                If gtot >= 0 Then
                    Dim wgtm As Double = 0
                    ' wgtm = Math.Round(((totalitemmean / titem) * myper), 2)
                    wgtm = Math.Round((totalitemmean / titem), 2)
                    row.Cells(1).Value = wgtm
                    averagewmean += wgtm
                    '    Addrow()
                    'If wgtm < 1.5 Then
                    '    row.Cells(2).Value = "Strongly Disagree"
                    'ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    '    row.Cells(2).Value = "Disagree"
                    'ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    '    row.Cells(2).Value = "Neutral"
                    'ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    '    row.Cells(2).Value = "Agree"
                    'ElseIf wgtm >= 4.5 Then
                    '    row.Cells(2).Value = "Strongly Agree"
                    'End If
                Else
                    row.Cells(1).Value = 0
                    averagewmean += 0
                    row.Cells(2).Value = 0
                End If

                row.Cells(0).Value = "Total Weighted Mean - " & gcat.Fields("partname").Value
                row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
                row.Cells(0).Style.BackColor = Color.LightGray
                Addrow()

                row = New DataGridViewRow()
                row.CreateCells(dg1)
                row.Cells(0).Style.BackColor = Color.Gray
                row.Cells(1).Style.BackColor = Color.Gray
                row.Cells(2).Style.BackColor = Color.Gray
                Addrow()
                pbar1.Value += myadd
                gcat.MoveNext()
            End While

            pbar1.Value = 0
            lblbarep.Text = ""
            cbosy.Enabled = True
            cboryear.Enabled = True
            cbocol.Enabled = True
            lblbarep.Visible = False
            pbar1.Visible = False
            btnexport.Enabled = True
            btnexport.Visible = True
        End If

        gcat.Close()



    End Sub

    Private Sub cbocol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbocol.SelectedIndexChanged

        If cbosy.Text <> Nothing And cboryear.Text <> Nothing Then
            Dim gc As New ADODB.Recordset
            sopen(gc)
            gc.Open("Select * from college where GroupName='" & cbocol.Text & "'", db)
            If gc.EOF = False Then
                ccol = gc.Fields("cno").Value
            End If
            gc.Close()

            cbosy.Enabled = False
            cboryear.Enabled = False
            cbocol.Enabled = False
            btnexport.Enabled = False

            setter()
        End If
    End Sub

    Private Sub btnexport_Click(sender As Object, e As EventArgs) Handles btnexport.Click
        'Try
        If dg1.Rows.Count > 0 Then


            Dim xlApp As Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet
            Dim misValue As Object = System.Reflection.Missing.Value

            Dim i As Int16, j As Int16

            xlApp = New Excel.ApplicationClass
            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = xlWorkBook.Sheets("sheet1")


            Dim cl As Integer = 0
            Dim mycol As Integer = 1
            For i = 0 To dg1.RowCount - 1
                cl = 0

                For j = 0 To dg1.ColumnCount - 1
                    cl += 1
                    xlWorkSheet.Cells(i + 1, cl) = dg1.Rows(i).Cells(j).Value
                    If Not IsNumeric(dg1.Rows(i).Cells(j).Value) Then
                        If dg1.Rows(i).Cells(j).Value = "School Year" Then
                            For x = 1 To dg1.ColumnCount
                                xlWorkSheet.Cells(i + 1, x).Font.Bold = True
                                xlWorkSheet.Cells(i + 1, x).Interior.ColorIndex = 6
                            Next
                        ElseIf dg1.Rows(i).Cells(j).Value = "Frequency" Or dg1.Rows(i).Cells(j).Value = "Items" Then
                            For x = 1 To dg1.ColumnCount
                                xlWorkSheet.Cells(i + 1, x).Font.Bold = True
                                xlWorkSheet.Cells(i + 1, x).Interior.ColorIndex = 4
                            Next
                        ElseIf dg1.Rows(i).Cells(j).Value = "TOTAL" Then
                            For x = 1 To dg1.ColumnCount
                                xlWorkSheet.Cells(i + 1, x).Font.Bold = True
                            Next
                        End If
                    End If

                Next
            Next
            xlWorkSheet.Range("A1:C600").EntireColumn.AutoFit()
            Dim myfilename As String = Nothing
            myfilename = "\" & cmenu & "_" & cbosy.Text & ".xls"

            'If IsNumeric(xlWorkSheet.Range("A1:C600")) Then
            'Else
            '    xlWorkSheet.Range("A1:C600").Interior.ColorIndex = 4
            '            'End If
            'End If
            Dim mpath As String = Nothing
            mpath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            xlWorkBook.SaveAs(mpath + myfilename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()

            Process.Start(mpath + myfilename)
         

            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)
        Else
            MsgBox("Insufficient Rows", 48, "ZERO ROWS!")
        End If
        'Catch ex As Exception
        '    MsgBox(ex)
        'End Try
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
            MessageBox.Show("Exception Occured while releasing object " + ex.ToString())
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Sub btncancelthread_Click(sender As Object, e As EventArgs) Handles btncancelthread.Click
        tr1.Abort()
        btnexport.Visible = True
        btnexport.Enabled = True
    End Sub
End Class