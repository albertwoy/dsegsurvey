Imports System.ComponentModel
Imports System.Threading
Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat

Public Class frmfrsreports
    Dim ripindex As Integer = 0
    Dim selyear As Integer = 0
    Dim ryear As Integer = 0
    Dim gtot As Integer = 0
    Dim ccol As Integer = 0
    Dim t1 As System.Threading.Thread
    Dim t2 As System.Threading.Thread
    Dim t3 As System.Threading.Thread
    Dim t4 As System.Threading.Thread
    Dim t5 As System.Threading.Thread
    Dim t6 As System.Threading.Thread
    Dim t7 As System.Threading.Thread
    Dim t8 As System.Threading.Thread
    Dim t9 As System.Threading.Thread
    Dim t10 As System.Threading.Thread
    Dim t11 As System.Threading.Thread
    Dim t12 As System.Threading.Thread
    Dim t13 As System.Threading.Thread

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
            gt.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
        Else
            gt.Open("select count(*) as cnt from respondent where myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            t1 = Nothing
            If t1 Is Nothing Then
                t1 = New Thread(Sub() getage())
                t1.IsBackground = True
                t1.Start()
            End If


            't1.Start(100)

            ' Start the asynchronous operation.
            'getage()
            ' getgender()
            'getreligion()
            'getotfam()
            'getbirthorder()
            'getmarrel()
            'getparext()
            'getplaceres()
            'getannualincome()
            'getinvolvement1()
            'getinvolvement2()
            'getsignificantothers()
            'getpart3()
            ' bgworker.RunWorkerAsync()
        End If

    End Sub

    Private Sub frmfrsreports_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        't1 = Nothing
        't2 = Nothing
        't3 = Nothing
        't4 = Nothing
        't5 = Nothing
        't6 = Nothing
        't7 = Nothing
        't8 = Nothing
        't9 = Nothing
        't10 = Nothing
        't11 = Nothing
        't12 = Nothing
        't13 = Nothing
        If Not t1 Is Nothing Then
            t1.Abort()
        End If
        If Not t2 Is Nothing Then
            t2.Abort()
        End If
        If Not t3 Is Nothing Then
            t3.Abort()
        End If
        If Not t4 Is Nothing Then
            t4.Abort()
        End If
        If Not t5 Is Nothing Then
            t5.Abort()
        End If
        If Not t6 Is Nothing Then
            t6.Abort()
        End If
        If Not t7 Is Nothing Then
            t7.Abort()
        End If
        If Not t8 Is Nothing Then
            t8.Abort()
        End If
        If Not t9 Is Nothing Then
            t9.Abort()
        End If
        If Not t10 Is Nothing Then
            t10.Abort()
        End If
        If Not t11 Is Nothing Then
            t11.Abort()
        End If
      
        If Not t12 Is Nothing Then
            t12.Abort()
        End If
        If Not t13 Is Nothing Then
            t13.Abort()
        End If
        't1
        't2.Abort()
        't3.Abort()
        't4.Abort()
        't5.Abort()
        't6.Abort()
        't7.Abort()
        't8.Abort()
        't9.Abort()
        't10.Abort()
        't11.Abort()
        't12.Abort()
        't13.Abort()


    End Sub

   

    Private Sub frmfrsreports_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
    
    Private Overloads Function AddRow() As Integer

        If InvokeRequired Then
            Return CInt(Invoke(New ReturnDelegate(AddressOf AddRow)))
        Else
            Return dg1.Rows.Add(row)
        End If

    End Function
    Sub getage()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)


            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Age ..."
            'Dim row As DataGridViewRow
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "School Year"
            row.Cells(1).Value = cbosy.Text
            row.Cells(2).Value = Nothing
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            Addrow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Age"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            ' Addrow()
            AddRow()
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "16-17"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=16 or age=17) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where (age=16 or age=17) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            AddRow() '  Addrow()
            grm.Close()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "18-19"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=18 or age=19) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where (age=18 or age=19) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            AddRow() 'Addrow()
            grm.Close()




            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "20-21"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=20 or age=21) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where (age=20 or age=21) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            AddRow() ' Addrow()
            grm.Close()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "22-23"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and (age=22 or age=23) and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where (age=22 or age=23) and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            AddRow() 'Addrow()
            grm.Close()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "24 & above"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and age>=24 and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where age>=24 and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            AddRow() ' Addrow()
            grm.Close()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            ' Addrow()
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            ' Addrow()
            AddRow()

            pbar1.Value += 5

            t2 = Nothing
            If t2 Is Nothing Then
                t2 = New Thread(Sub() getgender())
                t2.IsBackground = True
                t2.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try



    End Sub
    Sub getgender()

        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)


            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Gender ..."
            '  Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Gender"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Male"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and gender='Male' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where gender='Male' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Female"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and gender='Female' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where gender='Female' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            Addrow()

            pbar1.Value += 5

            t3 = Nothing
            If t3 Is Nothing Then
                t3 = New Thread(Sub() getreligion())
                t3.IsBackground = True
                t3.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
        ' Application.DoEvents()
        ' getreligion()

    End Sub

    Sub getreligion()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)


            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Religion ..."
            'Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Religion"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Born Again Christian"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Born Again Christian' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Born Again Christian' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Iglesia Ni Cristo"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Iglesia Ni Cristo' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Iglesia Ni Cristo' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Islam/Muslim"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Islam/Muslim' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Islam/Muslim' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Jehovahs Witness"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Jehovahs Witness' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Jehovahs Witness' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Later-Day Saints (Mormons)"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Later-Day Saints (Mormons)' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Later-Day Saints (Mormons)' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "PIC/IFI/Aglipay"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='PIC/IFI/Aglipay' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='PIC/IFI/Aglipay' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Protestant"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Protestant' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Protestant' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Roman Catholic"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Roman Catholic' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Roman Catholic' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "SDA"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='SDA' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='SDA' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Others"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and relg='Others' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where relg='Others' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4", db)
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
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)
            AddRow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            pbar1.Value += 10
            AddRow()


            t4 = Nothing
            If t4 Is Nothing Then
                t4 = New Thread(Sub() getotfam())
                t4.IsBackground = True
                t4.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
        '  Application.DoEvents()
        ' getotfam()

    End Sub
    Sub getotfam()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)

            ' Dim row As DataGridViewRow
            lblbarep.Text = "Please wait, Getting Frequency/Percentage for No. of Family Members ..."
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "No. of Family Members"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            For j = 8 To 2 Step -1


                row = New DataGridViewRow()
                row.CreateCells(dg1)
                If j = 8 Then
                    row.Cells(0).Value = j & " members & above"
                Else
                    row.Cells(0).Value = j
                End If


                If cbocol.Text <> Nothing Then
                    If j = 8 Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and totfam>=" & j & " and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and totfam=" & j & " and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    End If

                Else
                    If j = 8 Then
                        grm.Open("select count(*) as cnt from respondent where totfam>=" & j & " and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent where totfam=" & j & " and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
                    End If

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

            Next


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            AddRow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()

            pbar1.Value += 5

            t5 = Nothing
            If t5 Is Nothing Then
                t5 = New Thread(Sub() getbirthorder())
                t5.IsBackground = True
                t5.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
        'Application.DoEvents()
        'getbirthorder()

    End Sub
    Sub getbirthorder()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Birth Order ..."
            Dim grm As New ADODB.Recordset
            sopen(grm)

            '  Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Birth Order"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Eldest"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and youare='Eldest' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where youare='Eldest' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Middle"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and youare='Middle' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where youare='Middle' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Youngest"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and youare='Youngest' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where youare='Youngest' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Only Child"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and youare='Only Child' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where youare='Only Child' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            AddRow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()

            pbar1.Value += 5
            'Dim t6 As System.Threading.Thread
            t6 = Nothing
            If t6 Is Nothing Then
                t6 = New Thread(Sub() getmarrel())
                t6.IsBackground = True
                t6.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
    
        'Application.DoEvents()
        ' getmarrel()

    End Sub
    Sub getmarrel()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Marital Relationship ..."
            Dim grm As New ADODB.Recordset
            sopen(grm)

            ' Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Marital Relationship"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Living Together"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and youpar='Living Together' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where youpar='Living Together' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Separated/Annulled"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and youpar='Separated/Annulled' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where youpar='Separated/Annulled' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            AddRow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value += 5
            t7 = Nothing
            If t7 Is Nothing Then
                t7 = New Thread(Sub() getparext())
                t7.IsBackground = True
                t7.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
        'Application.DoEvents()
        'getparext()

    End Sub
    Sub getparext()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Parental Existence ..."
            ' Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Parental Existence"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Both Parents Alive"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and parlive='Yes' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where parlive='Yes' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Father Only Alive"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and parlive='No' and deceased='Mother' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where parlive='No' and deceased='Mother' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Mother Only Alive"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and parlive='No' and deceased='Father' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where parlive='No' and deceased='Father' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Both Deceased"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and parlive='No' and deceased='Both' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where parlive='No' and deceased='Both' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            AddRow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            pbar1.Value += 5
            AddRow()
            t8 = Nothing
            If t8 Is Nothing Then
                t8 = New Thread(Sub() getplaceres())
                t8.IsBackground = True
                t8.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
        'Application.DoEvents()
        'getplaceres()

    End Sub
    Sub getplaceres()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)


            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Place of Residence ..."
            'Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Place of Residence"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "City"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and livin='City' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where livin='City' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Town or Municipality"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and livin='Town or Municipality' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where livin='Town or Municipality' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "Barrio"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and livin='Barrio' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where livin='Barrio' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            AddRow()


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value += 5
            t9 = Nothing
            If t9 Is Nothing Then
                t9 = New Thread(Sub() getannualincome())
                t9.IsBackground = True
                t9.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
       
        'Application.DoEvents()
        'getannualincome()

    End Sub
    Sub getannualincome()
        Try
            Dim atot As Integer = 0
            Dim pertot As Double = 0

            Dim per As Double = 0


            Dim grm As New ADODB.Recordset
            sopen(grm)

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Family Annual Income ..."
            ' Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Family Annual Income"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "400,000 & above"
            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='400,000 & above' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='400,000 & above' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "300,000 - 400,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='300,000 - 400,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='300,000 - 400,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "270,000 - 300,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='270,000 - 300,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='270,000 - 300,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "240,000 - 270,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='240,000 - 270,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='240,000 - 270,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "210,000 - 240,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='210,000 - 240,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='210,000 - 240,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "180,000 - 210,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='180,000 - 210,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='180,000 - 210,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "150,000 - 180,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='150,000 - 180,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='150,000 - 180,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "120,000 - 150,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='120,000 - 150,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='120,000 - 150,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "90,000 - 120,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='90,000 - 120,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='90,000 - 120,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "60,000 - 90,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='60,000 - 90,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='60,000 - 90,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "30,000 - 60,000"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='30,000 - 60,000' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='30,000 - 60,000' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "30,000 and below"

            If cbocol.Text <> Nothing Then
                grm.Open("select count(distinct(r.rno)) as cnt from respondent r, courses c, college g where c.c_no=" & ccol & " and insrc='30,000 and below' and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
            Else
                grm.Open("select count(*) as cnt from respondent where insrc='30,000 and below' and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
            row.Cells(0).Value = "TOTAL"
            row.Cells(1).Value = atot
            row.Cells(2).Value = Math.Round(pertot)

            AddRow()

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value += 5
            t10 = Nothing
            If t10 Is Nothing Then
                t10 = New Thread(Sub() getinvolvement1())
                t10.IsBackground = True
                t10.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

        End Try
        'Application.DoEvents()
        'getinvolvement1()
        '50 percent
    End Sub
    Sub getinvolvement1()

        Try
            Dim per As Double = 0

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for INVOLVEMENT - non class related events ..."
            Dim grm As New ADODB.Recordset
            sopen(grm)

            'Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "INVOLVEMENT - non class related events"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            For i = 1 To 9
                row = New DataGridViewRow()
                row.CreateCells(dg1)
                row.Cells(0).Value = i

                If cbocol.Text <> Nothing Then
                    grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part41 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & " and rp1.rate='1' and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                Else
                    grm.Open("select count(*) as cnt from respondent r, respondent_part41 rp1 where rp1.item=" & i & " and rp1.rate='1' and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                AddRow()
                grm.Close()
            Next


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value += 5

            t11 = Nothing
            If t11 Is Nothing Then
                t11 = New Thread(Sub() getinvolvement2())
                t11.IsBackground = True
                t11.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
     
        'Application.DoEvents()
        'getinvolvement2()

    End Sub
    Sub getinvolvement2()
        Try
            Dim per As Double = 0

            Dim grm As New ADODB.Recordset
            sopen(grm)

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for INVOLVEMENT - classes related activities ..."
            'Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "INVOLVEMENT - classes related activities"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            For i = 1 To 11
                row = New DataGridViewRow()
                row.CreateCells(dg1)
                row.Cells(0).Value = i

                If cbocol.Text <> Nothing Then
                    grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part42 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & " and rp1.rate='1' and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                Else
                    grm.Open("select count(*) as cnt from respondent r, respondent_part42 rp1 where rp1.item=" & i & " and rp1.rate='1' and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                AddRow()
                grm.Close()
            Next

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value += 5
            t12 = Nothing
            If t12 Is Nothing Then
                t12 = New Thread(Sub() getsignificantothers())
                t12.IsBackground = True
                t12.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
       
        ' Application.DoEvents()
        ' getsignificantothers()

    End Sub
    Sub getsignificantothers()
        Try
            Dim per As Double = 0

            Dim grm As New ADODB.Recordset
            sopen(grm)

            'Dim row As DataGridViewRow

            lblbarep.Text = "Please wait, Getting Frequency/Percentage for Significant Others ..."
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Significant Others"
            row.Cells(1).Value = "Frequency"
            row.Cells(2).Value = "Percentage"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            ' DataGridView1.CurrentCell.Style.Font = New Font("Arial", 12, FontStyle.Bold)
            AddRow()

            For i = 1 To 7
                row = New DataGridViewRow()
                row.CreateCells(dg1)
                row.Cells(0).Value = i

                If cbocol.Text <> Nothing Then
                    grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part4 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.eventno=" & i & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                Else
                    grm.Open("select count(*) as cnt from respondent r, respondent_part4 rp1 where rp1.eventno=" & i & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                AddRow()
                grm.Close()
            Next


            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value += 5
            t13 = Nothing
            If t13 Is Nothing Then
                t13 = New Thread(Sub() getpart3())
                t13.IsBackground = True
                t13.Start()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
        End Try
       
        'Application.DoEvents()
        'getpart3()
    End Sub

    Sub getpart3()
        Try
            Dim per As Double = 0
            Dim averagewmean As Double = 0

            Dim grm As New ADODB.Recordset
            sopen(grm)

            'Dim row As DataGridViewRow

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Items"
            row.Cells(1).Value = "Weighted Mean"
            row.Cells(2).Value = "Interpretation"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            AddRow()

            lblbarep.Text = "Please wait... Computing Weighted Mean for ITEM - ACADEMIC ..."
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "ACADEMIC"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray


            Dim totitem() As Double
            Dim x As Integer
            ReDim Preserve totitem(8)
            x = 0
            For i = 7 To 14
                Application.DoEvents()
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            Dim totalitemmean As Double = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                '  wgtm = Math.Round(((totalitemmean / 8) * 0.15), 2)
                wgtm = Math.Round((totalitemmean / 8), 2)
                'row.Cells(1).Value = wgtm
                row.Cells(1).Value = wgtm
                averagewmean += wgtm

                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait... Computing Weighted Mean for ITEM - CAREER ..."
            pbar1.Value += 5
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "CAREER"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(6)
            x = 0
            For i = 15 To 20
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                'wgtm = Math.Round(((totalitemmean / 6) * 0.15), 2)
                wgtm = Math.Round((totalitemmean / 6), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait... Computing Weighted Mean for ITEM - SOCIAL ..."
            pbar1.Value += 5
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "SOCIAL"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(5)
            x = 0
            For i = 21 To 25
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                'wgtm = Math.Round(((totalitemmean / 5) * 0.1), 2)
                wgtm = Math.Round((totalitemmean / 5), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            '80 percent
            lblbarep.Text = "Please wait... Computing Weighted Mean for ITEM - INSTITUTONAL FACTORS ..."
            pbar1.Value += 5
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "INSTITUTONAL FACTORS"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(4)
            x = 0
            For i = 26 To 29
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                wgtm = Math.Round((totalitemmean / 4), 2)
                '   wgtm = Math.Round(((totalitemmean / 4) * 0.15), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait.... Computing Weighted Mean for ITEM - FACULTY ..."
            pbar1.Value += 5
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "FACULTY"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(5)
            x = 0
            For i = 30 To 34
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                'wgtm = Math.Round(((totalitemmean / 5) * 0.15), 2)
                wgtm = Math.Round((totalitemmean / 5), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait..... Computing Weighted Mean for ITEM - SUPPORT SERVICES ..."
            pbar1.Value += 5
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "SUPPORT SERVICES"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(10)
            x = 0
            For i = 35 To 44
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                'wgtm = Math.Round(((totalitemmean / 10) * 0.15), 2)
                wgtm = Math.Round((totalitemmean / 10), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait...... Computing Weighted Mean for ITEM - FINANCES ..."
            pbar1.Value += 5
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "FINANCES"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(7)
            x = 0
            For i = 45 To 51
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                ' wgtm = Math.Round(((totalitemmean / 7) * 0.1), 2)
                wgtm = Math.Round((totalitemmean / 7), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait....... Computing Weighted Mean for ITEM - PERSISTENCE ..."
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "PERSISTENCE"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(6)
            x = 0
            For i = 52 To 57
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                ' wgtm = Math.Round(((totalitemmean / 6) * 0.05), 2)
                wgtm = Math.Round((totalitemmean / 6), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()

            lblbarep.Text = "Please wait........ Computing Weighted Mean for ITEM - DIVERSITY ..."
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "DIVERSITY"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(3)
            x = 0
            For i = 58 To 60
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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

                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                'wgtm = Math.Round(((totalitemmean / 3) * 0.15), 2)
                wgtm = Math.Round((totalitemmean / 3), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If
            AddRow()


            lblbarep.Text = "Please wait......... Computing Weighted Mean for ITEM - PERSONAL ..."
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "PERSONAL"
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(0).Style.BackColor = Color.LightGray
            ReDim Preserve totitem(6)
            x = 0
            For i = 61 To 66
                Dim itemmean As Double = 0
                For j = 1 To 5
                    Dim tmpvar As Double = 0
                    If cbocol.Text <> Nothing Then
                        grm.Open("select count(distinct(r.rno)) as cnt from respondent r, respondent_part3 rp1, courses c, college g where c.c_no=" & ccol & " and  rp1.item=" & i & "  and rp1.rate=" & j & " and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and c.crse_no=r.course and surveyid=4", db)
                    Else
                        grm.Open("select count(*) as cnt from respondent r, respondent_part3 rp1 where rp1.item=" & i & " and rp1.rate=" & j & "  and r.rno=rp1.rno and myear=" & selyear & " and `year`=" & ryear & " and surveyid=4 ", db)
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
                '    x += 1
                If gtot > 0 Then
                    totitem(x) = Math.Round(itemmean / gtot, 2)
                Else
                    totitem(x) = 0
                End If
                x += 1
            Next

            totalitemmean = 0

            For x = 0 To UBound(totitem) - 1
                totalitemmean += totitem(x)
            Next
            If gtot >= 0 Then
                Dim wgtm As Double = 0
                'wgtm = Math.Round(((totalitemmean / 6) * 0.1), 2)
                wgtm = Math.Round((totalitemmean / 6), 2)
                row.Cells(1).Value = wgtm
                averagewmean += wgtm
                If wgtm < 1.5 Then
                    row.Cells(2).Value = "Strongly Disagree"
                ElseIf wgtm >= 1.5 And wgtm < 2.5 Then
                    row.Cells(2).Value = "Disagree"
                ElseIf wgtm >= 2.5 And wgtm < 3.5 Then
                    row.Cells(2).Value = "Neutral"
                ElseIf wgtm >= 3.5 And wgtm < 4.5 Then
                    row.Cells(2).Value = "Agree"
                ElseIf wgtm >= 4.5 Then
                    row.Cells(2).Value = "Strongly Agree"
                End If
            Else
                row.Cells(1).Value = 0
                averagewmean += 0
                row.Cells(2).Value = 0
            End If

            AddRow()


            averagewmean = Math.Round(averagewmean / 10, 2)
            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Value = "Average Weighted Mean"
            row.Cells(1).Value = averagewmean
            If averagewmean < 1.5 Then
                row.Cells(2).Value = "Strongly Disagree"
            ElseIf averagewmean >= 1.5 And averagewmean < 2.5 Then
                row.Cells(2).Value = "Disagree"
            ElseIf averagewmean >= 2.5 And averagewmean < 3.5 Then
                row.Cells(2).Value = "Neutral"
            ElseIf averagewmean >= 3.5 And averagewmean < 4.5 Then
                row.Cells(2).Value = "Agree"
            ElseIf averagewmean >= 4.5 Then
                row.Cells(2).Value = "Strongly Agree"
            End If
            'pbar1.Value += 5
            row.Cells(0).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(1).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            row.Cells(2).Style.Font = New Font("Arial", 10, FontStyle.Bold)
            AddRow()

            'pbar1.Value += 5

            row = New DataGridViewRow()
            row.CreateCells(dg1)
            row.Cells(0).Style.BackColor = Color.Gray
            row.Cells(1).Style.BackColor = Color.Gray
            row.Cells(2).Style.BackColor = Color.Gray

            AddRow()
            pbar1.Value = 0
            lblbarep.Text = ""
            cbosy.Enabled = True
            cboryear.Enabled = True
            cbocol.Enabled = True
            lblbarep.Visible = False
            pbar1.Visible = False

            btnexport.Enabled = True

            't1 = Nothing
            't2 = Nothing
            't3 = Nothing
            't4 = Nothing
            't5 = Nothing
            't6 = Nothing
            't7 = Nothing
            't8 = Nothing
            't9 = Nothing
            't10 = Nothing
            't11 = Nothing
            't12 = Nothing
            't13 = Nothing

        Catch ex As Exception

            MsgBox(ex.Message)

        Finally

        End Try
        

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

    
    Private Sub btncancelthread_Click(sender As Object, e As EventArgs) Handles btncancelthread.Click
        t1.Abort()
        btnexport.Visible = True
        btnexport.Enabled = True

    End Sub

    Private Sub btnexport_Click(sender As Object, e As EventArgs) Handles btnexport.Click
        Try
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
                myfilename = "/" & "Freshmen Retension Survey_" & cbosy.Text & ".xls"

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
                ' xlWorkBook = xlApp.Workbooks.Open(mpath + myfilename)
                'xlWorkSheet = xlWorkBook.Worksheets(1)

                releaseObject(xlWorkSheet)
                releaseObject(xlWorkBook)
                releaseObject(xlApp)
            Else
                MsgBox("Insufficient Rows", 48, "ZERO ROWS!")
            End If
        Catch ex As Exception
            MsgBox(ex)
        End Try
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
End Class