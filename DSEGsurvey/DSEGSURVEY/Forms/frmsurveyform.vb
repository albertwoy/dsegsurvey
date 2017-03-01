Public Class frmsurveyform
    Dim crow As Integer = 0
    Dim eventhandleradded As Boolean = False
    Dim mcsem As String = Nothing
    Dim mcyear As String = Nothing


    Sub displayq1()

        Dim qr As New ADODB.Recordset
        Dim row As New DataGridViewRow
        Dim qlim As Integer = 0

        Dim rcnt As Integer = 0
        qr.Open("Select * from tab_surveycategory where surveyid='" & svid & "'", db)
        While Not qr.EOF
            'rcnt += 1
            qlim = qr.Fields("itemcount").Value

            'If rcnt > 1 Then
            '    row = New DataGridViewRow()
            '    row.CreateCells(dgq1)
            '    dgq1.Rows.Add(row)
            'End If

            For i = 1 To qlim
                row = New DataGridViewRow()
                row.CreateCells(dgq1)
                row.Cells(0).Value = i
                row.Cells(1).Value = Nothing
                row.Cells(2).Value = qr.Fields("catid").Value
                dgq1.Rows.Add(row)

            Next

            qr.MoveNext()
        End While

        qr.Close()

        Dim cbCell As New DataGridViewComboBoxCell
        If svid = 1 Then
            For i = 0 To dgq1.Rows.Count - 1
                cbCell = dgq1.Rows(i).Cells(1)
                '  For iIndex = 0 To UBound(gSerialNumberArray)
                cbCell.Items.Add("5 - Very Much Needed")
                cbCell.Items.Add("4 - Much Needed")
                cbCell.Items.Add("3 - Somewhat Needed")
                cbCell.Items.Add("2 - Not Really Needed")
                cbCell.Items.Add("1 - Not at all")
                'Next
            Next
        ElseIf svid = 2 Then
            For i = 0 To dgq1.Rows.Count - 1
                cbCell = dgq1.Rows(i).Cells(1)
                '  For iIndex = 0 To UBound(gSerialNumberArray)
                cbCell.Items.Add("5 - Strongly Agree")
                cbCell.Items.Add("4 - Agree")
                cbCell.Items.Add("3 - Moderately Agree")
                cbCell.Items.Add("2 - Disagree")
                cbCell.Items.Add("1 - Strongly Disagree")
                'Next
            Next
        ElseIf svid = 3 Then
            For i = 0 To dgq1.Rows.Count - 1
                cbCell = dgq1.Rows(i).Cells(1)
                '  For iIndex = 0 To UBound(gSerialNumberArray)
                cbCell.Items.Add("5 - Strongly Agree")
                cbCell.Items.Add("4 - Agree")
                cbCell.Items.Add("3 - Somewhat Agree")
                cbCell.Items.Add("2 - Disagree")
                'Next
            Next
        End If



        For j = 0 To dgq1.ColumnCount - 1
            '  dgq1.Columns(j).Resizable = DataGridViewTriState.False
        Next

        For i = 0 To dgq1.RowCount - 1
            dgq1.Rows(i).Cells(0).ReadOnly = True
            dgq1.Rows(i).Cells(2).ReadOnly = True
            If dgq1.Rows(i).Cells(0).Value <> Nothing Then
                dgq1.Rows(i).Cells(1).ReadOnly = False
            Else
                dgq1.Rows(i).Cells(1).ReadOnly = True

            End If

            '   dgq1.Rows(i).Cells(2).ReadOnly = True

        Next
    End Sub


    Sub displayk1()

        Dim qr As New ADODB.Recordset
        Dim row As New DataGridViewRow

        'sopen(qr)
        'qr.Open("Select * from questions where CATCODE='" & lblselcat.Text & "'", db)
        'Dim qcnts As Integer = 0
        'For i = 5 To 0
        '    row = New DataGridViewRow()
        '    row.CreateCells(dgkey1)
        '    row.Cells(0).Value = i
        '    dgkey1.Rows.Add(row)
        'Next
        'For j = 0 To dgkey1.ColumnCount - 1
        '    'dgkey1.Columns(j).Resizable = DataGridViewTriState.False
        'Next
        For i = 0 To dgq1.RowCount - 1
            ' dgkey1.Rows(i).Cells(0).ReadOnly = True
            ' dgkey1.Rows(i).Cells(1).ReadOnly = False
            ' dgkey1.Rows(i).Cells(2).ReadOnly = True

        Next
    End Sub

    Sub getautoincrement()
        Dim gat As New ADODB.Recordset
        sopen(gat)
        gat.Open("SHOW TABLE STATUS WHERE `Name` = 'respondent'", db)
        tresno.Text = gat.Fields("Auto_increment").Value
        gat.Close()
    End Sub

    Private Sub frmsurveyform_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        UnregisterHotKey(Me.Handle, 300)
        UnregisterHotKey(Me.Handle, 400)
    End Sub

    
    Private Sub frmSurveyforms_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' tabforms_Click(1, e)

        RegisterHotKey(Me.Handle, 300, CTRL, Keys.R)
        RegisterHotKey(Me.Handle, 400, CTRL, Keys.S)

        mcsem = DSEGMAIN.tssem.Text
        mcyear = DSEGMAIN.tsyear.Text

        cbocrs1.Items.Clear()
        Dim gc As New ADODB.Recordset
        gc.Open("Select * from courses where c_no<>'10' and c_no<>'11' and c_no<>'12'  and  c_no<>'14' order by crse_no", db)
        While Not gc.EOF
            cbocrs1.Items.Add(gc.Fields("crse_no").Value)
            gc.MoveNext()
        End While
        gc.Close()

        getautoincrement()

        If svid = 1 Then
            lblsurveyid.Text = 1
            dgq1.Rows.Clear()
            'dgkey1.Rows.Clear()
            displayq1()
            displayk1()
        ElseIf svid = 2 Then
            lblsurveyid.Text = 2
            dgq1.Rows.Clear()
            ' dgkey1.Rows.Clear()
            displayq1()
            displayk1()
        ElseIf svid = 3 Then
            lblsurveyid.Text = 3
            dgq1.Rows.Clear()
            'dgkey1.Rows.Clear()
            displayq1()
            displayk1()
        End If
      
       

    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                'Case "100"
                '   MessageBox.Show("You pressed ALT+D key combination")
                'Case "200"
                '   MessageBox.Show("You pressed ALT+C key combination")
                Case "300"
                    tsreset.PerformClick()
                Case "400"
                    saveme.PerformClick()
            End Select
        End If
        MyBase.WndProc(m)
    End Sub
    Private Sub dgq1_CellEnter(sender As Object, e As DataGridViewCellEventArgs)
        If e.ColumnIndex = 2 Then
            dgq1.BeginEdit(True)
        End If
        crow = e.RowIndex
        lblcrow.Text = crow
        dgq1.EditMode = DataGridViewEditMode.EditOnEnter

    End Sub

    Private Sub dgq1_KeyDown(sender As Object, e As KeyEventArgs)
        ' MsgBox(e.KeyCode)

        If e.KeyCode = Keys.Up Then
            If crow < dgq1.RowCount - 1 Then
                dgq1.Rows(crow + 1).Cells(1).Selected = True
            End If

        ElseIf e.KeyCode = Keys.Down Then
            If crow > 0 Then
                dgq1.Rows(crow - 1).Cells(1).Selected = True
            End If
        ElseIf e.KeyCode = Keys.F2 Then
            ' e.Handled = False
            dgq1.CurrentCell = dgq1.Rows(crow).Cells(1)
            dgq1.BeginEdit(True)
            dgq1.Focus()

        End If

    End Sub

    Private Sub saveme_Click(sender As Object, e As EventArgs)
        MsgBox("test")
    End Sub


    Private Sub dgq1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs)
        If eventhandleradded = False Then
            AddHandler e.Control.KeyDown, AddressOf cell_Keydown
            eventhandleradded = True
        End If
    End Sub
    Private Sub cell_Keydown(sender As Object, e As KeyEventArgs)
        If e.Control AndAlso e.KeyCode = 83 Then
            e.Handled = True
            e.SuppressKeyPress = True
            saveme.PerformClick()
        End If
    End Sub

    Private Sub tsreset_Click(sender As Object, e As EventArgs) Handles tsreset.Click
        For i = 0 To dgq1.Rows.Count - 1
            dgq1.Rows(i).Cells(1).Value = Nothing

        Next
       
        'lbl()
       

        dgq1.CurrentCell = dgq1.Rows(0).Cells(1)
        dgq1.Focus()

        cboyr1.Text = Nothing
        cbocrs1.Text = Nothing
        cbogender.Text = Nothing
        txtage.Text = Nothing
        txthome.Text = Nothing

        cboyr1.Enabled = True
        cbocrs1.Enabled = True
        cbogender.Enabled = True
        txtage.Enabled = True
        txthome.Enabled = True
        dgq1.Enabled = True

        lblrno.Text = Nothing

        cbocrs1.Focus()

        lblrno.Text = Nothing
        lblrn.Visible = False
        lblrno.Visible = False
        For i = 0 To dgq1.Rows.Count - 1
            dgq1.Rows(i).Cells(1).Value = Nothing

        Next

        getautoincrement()

    End Sub

    Private Sub saveme_Click_1(sender As Object, e As EventArgs) Handles saveme.Click
        Dim c As Integer = 0
        If MsgBox("Are you sure you want to save this survey?", 4 + 32, "Saved Survey?") = 7 Then
            c = 1
        Else
            If cbocrs1.Text = Nothing Or cboyr1.Text = Nothing Then
                MsgBox("Please select course and year to continue", 48, "PLEASE SELECT")
                Exit Sub
            End If

            Dim myrows As Integer = 0
            myrows = dgq1.Rows.Count - 1

            For i = 0 To dgq1.Rows.Count - 1

                If i = myrows Then
                    If dgq1.Rows(myrows).Cells(1).Value = Nothing Then
                        MsgBox("Please press enter key on last row before saving", 48, "PLEASE PRESS ENTER KEY")
                        Exit Sub
                        Exit For

                    End If
                End If
            Next

            If svid = 1 Then
                saveneeda()
            ElseIf svid = 2 Then
                savecampusenvironmenta()
            ElseIf svid = 3 Then
                savevaluationguidance()
            End If

            '    saverating()

            cbocrs1.Text = Nothing
            cboyr1.Text = Nothing
            cbogender.Text = Nothing
            txtage.Text = Nothing
            txthome.Text = Nothing
            For i = 0 To dgq1.Rows.Count - 1
                dgq1.Rows(i).Cells(1).Value = Nothing
            Next
            cbocrs1.Focus()
            MsgBox("Record Successfully Saved", 64, "RECORD SAVED!")
        End If
       
    End Sub
    Sub saveneeda()
        Dim myuser As String = Nothing
        Dim rnno As String = Nothing

        myuser = DSEGMAIN.tsstatus.Text

        Dim sn As New ADODB.Recordset
        sopen(sn)
        If lblrn.Visible = True And lblrno.Text <> Nothing Then
            rnno = lblrno.Text
            db.Execute("UPDATE respondent SET course='" & cbocrs1.Text & "',`year`='" & cboyr1.Text & "',gender='" & cbogender.Text & "'" & _
                       ",homeaddress='" & txthome.Text & "',datesaved='" & getdatetime(1) & "' where rno='" & rnno & "' and surveyid='" & svid & "'")
        Else
            sn.Open("Select * from respondent", db)

            sn.AddNew()
            sn.Fields("course").Value = cbocrs1.Text
            sn.Fields("year").Value = cboyr1.Text
            sn.Fields("gender").Value = cbogender.Text
            sn.Fields("homeaddress").Value = txthome.Text
            If txtage.Text <> Nothing Then
                sn.Fields("age").Value = txtage.Text
            End If

            sn.Fields("surveyid").Value = svid
            sn.Fields("msem").Value = mcsem
            sn.Fields("myear").Value = mcyear
            sn.Fields("datesaved").Value = getdatetime(1)
            sn.Fields("savedby").Value = myuser
            sn.Update()
            sn.Close()



            sn.Open("Select last_insert_id() as newrno from respondent", db)
            tresno.Text = sn.Fields("newrno").Value
            rnno = sn.Fields("newrno").Value
            sn.Close()
            lblrn.Visible = False
            lblrno.Text = Nothing

        End If
        

      

        For i = 0 To dgq1.Rows.Count - 1

            If dgq1.Rows(i).Cells(1).Value <> Nothing Then

                sn.Open("Select * from tab_surveyans where rno='" & rnno & "' and qid='" & dgq1.Rows(i).Cells(0).Value & "' and catid='" & dgq1.Rows(i).Cells(2).Value & "'", db)
                If sn.EOF = True Then

                    sn.AddNew()
                    sn.Fields("rno").Value = tresno.Text
                    sn.Fields("qid").Value = dgq1.Rows(i).Cells(0).Value
                    sn.Fields("rate").Value = Strings.Left(dgq1.Rows(i).Cells(1).Value, 1)
                    sn.Fields("catid").Value = dgq1.Rows(i).Cells(2).Value
                    sn.Fields("surveyid").Value = svid
                    sn.Fields("msem").Value = mcsem
                    sn.Fields("myear").Value = mcyear
                    sn.Fields("mydate").Value = getdatetime(1)
                    sn.Update()

                ElseIf sn.EOF = False Then

                    db.Execute("UPDATE tab_surveyans SET  rate='" & Strings.Left(dgq1.Rows(i).Cells(1).Value, 1) & "', mydate='" & getdatetime(1) & "' where rno='" & rnno & "' and qid='" & dgq1.Rows(i).Cells(0).Value & "' and catid='" & dgq1.Rows(i).Cells(2).Value & "' and surveyid='" & svid & "'")
                    'sn.Fields("qid").Value = dgq1.Rows(i).Cells(0).Value
                    'sn.Fields("rate").Value = Strings.Left(dgq1.Rows(i).Cells(1).Value, 1)
                    'sn.Fields("catid").Value = dgq1.Rows(i).Cells(2).Value
                    'sn.Fields("surveyid").Value = svid
                    'sn.Fields("msem").Value = mcsem
                    'sn.Fields("myear").Value = mcyear
                    'sn.Fields("mydate").Value = getdatetime(1)

                    'sn.Update()
                End If
                sn.Close()
            End If
        Next
        lblrno.Text = Nothing
        lblrn.Visible = False
        lblrno.Visible = False
    End Sub
    Sub savecampusenvironmenta()
        Dim myuser As String = Nothing
        Dim rnno As String = Nothing

        myuser = DSEGMAIN.tsstatus.Text

        Dim sc As New ADODB.Recordset
        sopen(sc)
        If lblrn.Visible = True And lblrno.Text <> Nothing Then
            rnno = lblrno.Text
            db.Execute("UPDATE respondent SET course='" & cbocrs1.Text & "',`year`='" & cboyr1.Text & "',gender='" & cbogender.Text & "'" & _
                       ",homeaddress='" & txthome.Text & "',datesaved='" & getdatetime(1) & "' where rno='" & rnno & "' and surveyid='" & svid & "'")
        Else
            sc.Open("Select * from respondent", db)

            sc.AddNew()
            sc.Fields("course").Value = cbocrs1.Text
            sc.Fields("year").Value = cboyr1.Text
            sc.Fields("gender").Value = cbogender.Text
            sc.Fields("homeaddress").Value = txthome.Text
            If txtage.Text <> Nothing Then
                sc.Fields("age").Value = txtage.Text
            End If
            sc.Fields("surveyid").Value = svid
            sc.Fields("msem").Value = mcsem
            sc.Fields("myear").Value = mcyear
            sc.Fields("datesaved").Value = getdatetime(1)
            sc.Fields("savedby").Value = myuser
            sc.Update()
            sc.Close()


            sc.Open("Select last_insert_id() as newrno from respondent", db)
            tresno.Text = sc.Fields("newrno").Value
            rnno = sc.Fields("newrno").Value
            sc.Close()
        End If



        For i = 0 To dgq1.Rows.Count - 1
            If dgq1.Rows(i).Cells(1).Value <> Nothing Then

                sc.Open("Select * from tab_surveyans where rno='" & rnno & "' and qid='" & dgq1.Rows(i).Cells(0).Value & "' and catid='" & dgq1.Rows(i).Cells(2).Value & "'", db)
                If sc.EOF = True Then

                    sc.AddNew()
                    sc.Fields("rno").Value = tresno.Text
                    sc.Fields("qid").Value = dgq1.Rows(i).Cells(0).Value
                    sc.Fields("rate").Value = Strings.Left(dgq1.Rows(i).Cells(1).Value, 1)
                    sc.Fields("catid").Value = dgq1.Rows(i).Cells(2).Value
                    sc.Fields("surveyid").Value = svid
                    sc.Fields("msem").Value = mcsem
                    sc.Fields("myear").Value = mcyear
                    sc.Fields("mydate").Value = getdatetime(1)
                    sc.Update()

                ElseIf sc.EOF = False Then

                    db.Execute("UPDATE tab_surveyans SET  rate='" & Strings.Left(dgq1.Rows(i).Cells(1).Value, 1) & "', mydate='" & getdatetime(1) & "'  where rno='" & rnno & "' and qid='" & dgq1.Rows(i).Cells(0).Value & "' and catid='" & dgq1.Rows(i).Cells(2).Value & "' and surveyid='" & svid & "'")
                    'sc.Fields("qid").Value = dgq1.Rows(i).Cells(0).Value
                    'sc.Fields("rate").Value = Strings.Left(dgq1.Rows(i).Cells(1).Value, 1)
                    'sc.Fields("catid").Value = dgq1.Rows(i).Cells(2).Value
                    'sc.Fields("surveyid").Value = svid
                    'sc.Fields("msem").Value = mcsem
                    'sc.Fields("myear").Value = mcyear
                    'sc.Fields("mydate").Value = getdatetime(1)

                    'sc.Update()
                End If
                sc.Close()
            End If
        Next
        lblrno.Text = Nothing
        lblrn.Visible = False
        lblrno.Visible = False
    End Sub
    Sub savevaluationguidance()
        Dim myuser As String = Nothing
        myuser = DSEGMAIN.tsstatus.Text
        Dim rnno As String = Nothing


        Dim svg As New ADODB.Recordset
        sopen(svg)

        If lblrn.Visible = True And lblrno.Text <> Nothing Then
            rnno = lblrno.Text
            db.Execute("UPDATE respondent SET course='" & cbocrs1.Text & "',`year`='" & cboyr1.Text & "',gender='" & cbogender.Text & "'" & _
                       ",homeaddress='" & txthome.Text & "',datesaved='" & getdatetime(1) & "' where rno='" & rnno & "' and surveyid='" & svid & "'")
        Else
            svg.Open("Select * from respondent", db)

            svg.AddNew()
            svg.Fields("course").Value = cbocrs1.Text
            svg.Fields("year").Value = cboyr1.Text
            svg.Fields("gender").Value = cbogender.Text
            svg.Fields("homeaddress").Value = txthome.Text
            If txtage.Text <> Nothing Then
                svg.Fields("age").Value = txtage.Text
            End If
            svg.Fields("surveyid").Value = svid
            svg.Fields("msem").Value = mcsem
            svg.Fields("myear").Value = mcyear
            svg.Fields("datesaved").Value = getdatetime(1)
            svg.Fields("savedby").Value = myuser
            svg.Update()
            svg.Close()

            svg.Open("Select last_insert_id() as newrno from respondent", db)
            tresno.Text = svg.Fields("newrno").Value
            rnno = svg.Fields("newrno").Value
            svg.Close()

        End If
        
      
        For i = 0 To dgq1.Rows.Count - 1
            If dgq1.Rows(i).Cells(1).Value <> Nothing Then

                svg.Open("Select * from tab_surveyans where rno='" & rnno & "' and qid='" & dgq1.Rows(i).Cells(0).Value & "' and catid='" & dgq1.Rows(i).Cells(2).Value & "'", db)
                If svg.EOF = True Then

                    svg.AddNew()
                    svg.Fields("rno").Value = tresno.Text
                    svg.Fields("qid").Value = dgq1.Rows(i).Cells(0).Value
                    svg.Fields("rate").Value = Strings.Left(dgq1.Rows(i).Cells(1).Value, 1)
                    svg.Fields("catid").Value = dgq1.Rows(i).Cells(2).Value
                    svg.Fields("surveyid").Value = svid
                    svg.Fields("msem").Value = mcsem
                    svg.Fields("myear").Value = mcyear
                    svg.Fields("mydate").Value = getdatetime(1)
                    svg.Update()

                ElseIf svg.EOF = False Then
                    db.Execute("UPDATE tab_surveyans SET  rate='" & Strings.Left(dgq1.Rows(i).Cells(1).Value, 1) & "', mydate='" & getdatetime(1) & "'  where rno='" & rnno & "' and qid='" & dgq1.Rows(i).Cells(0).Value & "' and catid='" & dgq1.Rows(i).Cells(2).Value & "' and surveyid='" & svid & "' ")
                    'svg.Fields("qid").Value = dgq1.Rows(i).Cells(0).Value
                    'svg.Fields("rate").Value = Strings.Left(dgq1.Rows(i).Cells(1).Value, 1)
                    'svg.Fields("catid").Value = dgq1.Rows(i).Cells(2).Value
                    'svg.Fields("surveyid").Value = svid
                    'svg.Fields("msem").Value = mcsem
                    'svg.Fields("myear").Value = mcyear
                    'svg.Fields("mydate").Value = getdatetime(1)

                    'svg.Update()
                End If
                svg.Close()
            End If
        Next
        lblrno.Text = Nothing
        lblrn.Visible = False
        lblrno.Visible = False
    End Sub

    Sub saverating()
        Dim svr As New ADODB.Recordset
        sopen(svr)


        For j = 0 To dgq1.Rows.Count - 1
            svr.Open("Select * from tab_surveyans", db)
            svr.AddNew()
            svr.Fields("rno").Value = tresno.Text
            svr.Fields("catid").Value = dgq1.Rows(j).Cells(2).Value
            svr.Fields("qid").Value = dgq1.Rows(j).Cells(0).Value
            svr.Fields("surveyid").Value = svid
            svr.Fields("rate").Value = Strings.Left(dgq1.Rows(j).Cells(1).Value, 1)
            svr.Fields("msem").Value = mcsem
            svr.Fields("myear").Value = mcyear
            svr.Fields("mydate").Value = getdatetime(1)
            svr.Update()
            svr.Close()
        Next

    End Sub


    Private Sub cmdclear_Click(sender As Object, e As EventArgs)
        cbocrs1.Text = Nothing
        cboyr1.Text = Nothing
        cbogender.Text = Nothing
        txtage.Text = Nothing
        txthome.Text = Nothing
        For i = 0 To dgq1.Rows.Count - 1
            dgq1.Rows(i).Cells(1).Value = Nothing
        Next
    End Sub

    Private Sub tsrefresh_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub tsreport_Click(sender As Object, e As EventArgs) Handles tsreport.Click

        With frmsurveyreport
            .MdiParent = DSEGMAIN
            ' .Dock = DockStyle.Fill
            .BringToFront()
            .Show()
        End With
    End Sub

   

    Private Sub tsrespondent_Click(sender As Object, e As EventArgs) Handles tsrespondent.Click
        With frmrespondents
            .lblsno.Text = lblsurveyid.Text
            .Show(Me)
        End With

    End Sub

    
    Private Sub cbocrs1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbocrs1.SelectedIndexChanged

    End Sub
End Class
