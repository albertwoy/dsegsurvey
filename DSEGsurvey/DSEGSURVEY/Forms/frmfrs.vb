
Public Class frmfrs
    Dim msm As String = Nothing
    Dim mym As String = Nothing
    Dim ps1 As New ADODB.Recordset

    Private Sub frmfrs_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        UnregisterHotKey(Me.Handle, 300)
        UnregisterHotKey(Me.Handle, 400)
    End Sub

    Private Sub frmfrs_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then
            saveme.PerformClick()
        ElseIf (e.KeyCode = Keys.R AndAlso e.Modifiers = Keys.Control) Then
            tsreset.PerformClick()
        End If
    End Sub

    Private Sub frmfrs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' RegisterHotKey(Me.Handle, 100, MOD_ALT, Keys.D)
        'RegisterHotKey(Me.Handle, 200, MOD_ALT, Keys.C)
        RegisterHotKey(Me.Handle, 300, CTRL, Keys.R)
        RegisterHotKey(Me.Handle, 400, CTRL, Keys.S)

        clearer()
        getautoincrement()

        msm = DSEGMAIN.tssem.Text
        mym = DSEGMAIN.tsyear.Text
        Dim gc As New ADODB.Recordset
        gc.Open("Select * from courses where c_no<>'10' and c_no<>'11' and c_no<>'12' and  c_no<>'14' order by crse_no", db)
        While Not gc.EOF
            t2.Items.Add(gc.Fields("crse_no").Value)
            gc.MoveNext()
        End While
        gc.Close()
        populategrid()
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
    Sub populategrid()
        'part3
        Dim row As New DataGridViewRow
        dg2.Rows.Clear()
        For i = 7 To 66
            row = New DataGridViewRow()
            row.CreateCells(dg2)
            row.Cells(0).Value = i
            row.Cells(1).Value = Nothing
            dg2.Rows.Add(row)
        Next

        dg3.Rows.Clear()
        For i = 1 To 9
            row = New DataGridViewRow()
            row.CreateCells(dg3)
            row.Cells(0).Value = i
            row.Cells(1).Value = Nothing
            dg3.Rows.Add(row)
        Next

        dg4.Rows.Clear()
        For i = 1 To 11
            row = New DataGridViewRow()
            row.CreateCells(dg4)
            row.Cells(0).Value = i
            row.Cells(1).Value = Nothing
            dg4.Rows.Add(row)
        Next

        For i = 0 To dg2.Rows.Count - 1
            dg2.Rows(i).Cells(0).ReadOnly = True
        Next
        For i = 0 To dg3.Rows.Count - 1
            dg3.Rows(i).Cells(0).ReadOnly = True
        Next
        For i = 0 To dg4.Rows.Count - 1
            dg4.Rows(i).Cells(0).ReadOnly = True
        Next


        'part3 -------------END
    End Sub
    Private Sub saveme_Click(sender As Object, e As EventArgs) Handles saveme.Click
        If t2.Text = Nothing Or t3.Text = Nothing Then
            MsgBox("Please select course and year to continue", 48, "PLEASE SELECT")
            tcr.SelectedIndex = 0
            Exit Sub
        End If
        If t13.Text = "No" And t14.Text = Nothing Then
            MsgBox("Please select deceased field to continue", 48, "PLEASE SELECT")
            Exit Sub
        End If
        If t17.Text = "Yes" And t18.Text = Nothing Then
            MsgBox("Please select kind of abused field to continue", 48, "PLEASE SELECT")
            Exit Sub
        End If
        If t19.Text = "Yes" And t20.Text = Nothing Then
            MsgBox("Please select input neglected details to continue", 48, "PLEASE SELECT")
            Exit Sub
        End If

        If t13.Text = "No" And t14.Text = Nothing Then
            MsgBox("Please select deceased to continue", 48, "PLEASE SELECT")
            Exit Sub
        End If

        'Dim myrows As Integer = 0
        'myrows = dg2.Rows.Count - 1

        'For i = 0 To dg2.Rows.Count - 1

        '    If i = myrows Then
        '        If dg2.Rows(myrows).Cells(1).Value = Nothing Then
        '            MsgBox("Please press enter key ON [PART 3 SECTION] at last row before saving", 48, "PLEASE PRESS ENTER KEY")
        '            Exit Sub
        '            Exit For

        '        End If
        '    End If
        'Next

        'myrows = dg3.Rows.Count - 1
        'For i = 0 To dg3.Rows.Count - 1

        '    If i = myrows Then
        '        If dg3.Rows(myrows).Cells(1).Value = Nothing Then
        '            MsgBox("Please press enter key ON [PART 4 RELATED EVENTS SECTION] at last row before saving", 48, "PLEASE PRESS ENTER KEY")
        '            Exit Sub
        '            Exit For

        '        End If
        '    End If
        'Next

        'myrows = dg4.Rows.Count - 1
        'For i = 0 To dg4.Rows.Count - 1

        '    If i = myrows Then
        '        If dg4.Rows(myrows).Cells(1).Value = Nothing Then
        '            MsgBox("Please press enter key ON [PART 4 CLASSES RELATED SECTION] at last row before saving", 48, "PLEASE PRESS ENTER KEY")
        '            Exit Sub
        '            Exit For

        '        End If
        '    End If
        'Next


        Dim c As Integer = 0
        If MsgBox("Are you sure you want to save this survey?", 4 + 32, "Save Survey") = 7 Then
            c = 1
        Else
            savepart1()
            savepart21()
            savepart22()
            savepart3()
            savepart4()
            MsgBox("Record Successfully Saved", 64, "RECORD SAVED!")
            getautoincrement()
            populategrid()

            clearer()
            tcr.SelectedIndex = 0

        End If


    End Sub

    Sub savepart1()


        sopen(ps1)
        If lblrn.Visible = True And lblrno.Text <> Nothing Then
            ps1.Open("Select * from respondent where rno='" & lblrno.Text & "'", db)
            If ps1.EOF = False Then
                savep1()
            End If
            tresno.Text = lblrno.Text
            ps1.Close()
        Else

            ps1.Open("Select * from respondent", db)
            ps1.AddNew()
            savep1()
            ps1.Close()

            ps1.Open("Select last_insert_id() as newrno from respondent", db)
            tresno.Text = ps1.Fields("newrno").Value
            ps1.Close()
        End If


    End Sub
    Sub savep1()
        ps1.Fields("rname").Value = t1.Text
        ps1.Fields("course").Value = t2.Text
        ps1.Fields("year").Value = t3.Text
        ps1.Fields("gender").Value = t4.Text
        If t5.Text <> Nothing Then
            ps1.Fields("age").Value = t5.Text
        End If
        ps1.Fields("wgt").Value = t6.Text
        ps1.Fields("hgt").Value = t7.Text
        ps1.Fields("relg").Value = t8.Text
        If t9.Text <> Nothing Then
            ps1.Fields("totsib").Value = t9.Text
        End If
        If t10.Text <> Nothing Then
            ps1.Fields("totfam").Value = t10.Text
        End If

        ps1.Fields("youare").Value = t11.Text
        ps1.Fields("youpar").Value = t12.Text
        ps1.Fields("parlive").Value = t13.Text
        ps1.Fields("deceased").Value = t14.Text

        ps1.Fields("livin").Value = t15.Text
        ps1.Fields("insrc").Value = t16.Text
        ps1.Fields("abused").Value = t17.Text
        ps1.Fields("abusedy").Value = t18.Text
        ps1.Fields("negl").Value = t19.Text

        ps1.Fields("negly").Value = t20.Text
        Dim myasm As String = Nothing

        If t38.Checked = True Then
            myasm = t38.Text
        ElseIf t39.Checked = True Then
            myasm = t39.Text
        ElseIf t40.Checked = True Then
            myasm = t40.Text
        ElseIf t41.Checked = True Then
            myasm = t41.Text
        ElseIf t42.Checked = True Then
            myasm = t42.Text
        ElseIf t43.Checked = True Then
            myasm = t43.Text
        End If

        If myasm <> Nothing Then
            ps1.Fields("studmu").Value = myasm
        End If

        Dim myasm2 As String = Nothing
        If t44.Checked = True Then
            myasm2 = t44.Text
        ElseIf t45.Checked = True Then
            myasm2 = t45.Text
        ElseIf t46.Checked = True Then
            myasm2 = t46.Text
        ElseIf t47.Checked = True Then
            myasm2 = t47.Text
        End If
        If myasm2 <> Nothing Then
            ps1.Fields("lastsem").Value = myasm2
        End If

        If t48.Text <> Nothing Then
            ps1.Fields("nummet").Value = t48.Text
        End If
        If t49.Text <> Nothing Then
            ps1.Fields("numclass").Value = t49.Text
        End If

        ps1.Fields("anything").Value = t80.Text
        ps1.Fields("msem").Value = msm
        ps1.Fields("myear").Value = mym
        Dim msid As Integer = 0
        msid = lblsurveyid.Text

        ps1.Fields("surveyid").Value = msid
        ps1.Fields("datesaved").Value = getdatetime(1)
        Dim myuser As String = Nothing
        myuser = DSEGMAIN.tsstatus.Text
        ps1.Fields("savedby").Value = myuser
        ps1.Update()




    End Sub

    Sub savepart21()


        Dim mydet As String = Nothing

        Dim TeamIndex(), i As Integer

        Dim x As Integer = 0
        ReDim Preserve TeamIndex(8)
        If t21.Checked = True Then
            TeamIndex(x) = t21.Text
            x += 1
        End If
        If t22.Checked = True Then
            TeamIndex(x) = t22.Text
            x += 1
        End If
        If t23.Checked = True Then
            TeamIndex(x) = t23.Text
            x += 1
        End If
        If t24.Checked = True Then
            TeamIndex(x) = t24.Text
            x += 1
        End If
        If t25.Checked = True Then
            TeamIndex(x) = t25.Text
            x += 1
        End If
        If t26.Checked = True Then
            TeamIndex(x) = t26.Text
            x += 1
        End If
        If t27.Checked = True Then
            TeamIndex(x) = t27.Text
            x += 1
        End If
        If t28.Checked = True Then
            TeamIndex(x) = t28.Text
            x += 1
        End If
        x = x - 1
        ReDim Preserve TeamIndex(x)

        If t28.Checked = True Then

            If t29.Text <> Nothing Then
                mydet = t29.Text
            Else
                mydet = Nothing
            End If
        End If


        Dim ps21 As New ADODB.Recordset
        sopen(ps21)
        Dim mk As Integer = 0
        For i = 0 To UBound(TeamIndex)
            '   MsgBox(TeamIndex(i))
            If mk = 0 Then
                db.Execute("DELETE from respondent_part21 where rno='" & tresno.Text & "'")
                mk = 1
            End If

            ps21.Open("Select * from respondent_part21 where rno='" & tresno.Text & "' and reason='" & TeamIndex(i) & "'", db)
            If ps21.EOF = True Then

                ps21.AddNew()
                ps21.Fields("rno").Value = tresno.Text
                ps21.Fields("reason").Value = TeamIndex(i)
                If TeamIndex(i) = 8 Then
                    ps21.Fields("details").Value = mydet
                End If

                ps21.Update()

            ElseIf ps21.EOF = False Then
                'MsgBox(TeamIndex(i))
                ' ps21.Fields("reason").Value = TeamIndex(i)
                'ps21.Fields("details").Value = mydet
                ' ps21.Update()


                ' db.Execute("INSERT INTO respondent_part21(rno,reason,details) VALUES('" & tresno.Text & "','" & TeamIndex(i) & "','" & mydet & "')")
                'db.Execute("UPDATE respondent_part21 SET details='" & mydet & "' where rno='" & tresno.Text & "' and reason='" & TeamIndex(i) & "'")
            End If
            ps21.Close()
        Next
    End Sub
    Sub savepart22()
        Dim mydet As String = Nothing

        Dim TeamIndex(), i As Integer

        Dim x As Integer = 0
        ReDim Preserve TeamIndex(8)
        If t30.Checked = True Then
            TeamIndex(x) = t30.Text
            x += 1
        End If
        If t31.Checked = True Then
            TeamIndex(x) = t31.Text
            x += 1
        End If
        If t32.Checked = True Then
            TeamIndex(x) = t32.Text
            x += 1
        End If
        If t33.Checked = True Then
            TeamIndex(x) = t33.Text
            x += 1
        End If
        If t34.Checked = True Then
            TeamIndex(x) = t34.Text
            x += 1
        End If
        If t35.Checked = True Then
            TeamIndex(x) = t35.Text
            x += 1
        End If
        If t36.Checked = True Then
            TeamIndex(x) = t36.Text
            x += 1
        End If
        x = x - 1
        ReDim Preserve TeamIndex(x)

        If t36.Checked = True Then

            If t36.Text <> Nothing Then
                mydet = t37.Text
            Else
                mydet = Nothing
            End If
        End If


        Dim ps22 As New ADODB.Recordset
        sopen(ps22)
        Dim mk As Integer = 0

        For i = 0 To UBound(TeamIndex)
            '   MsgBox(TeamIndex(i))
            If mk = 0 Then
                db.Execute("DELETE from respondent_part22 where rno='" & tresno.Text & "'")
                mk = 1
            End If

            ps22.Open("Select * from respondent_part22 where rno='" & tresno.Text & "' and live='" & TeamIndex(i) & "'", db)
            If ps22.EOF = True Then

                ps22.AddNew()
                ps22.Fields("rno").Value = tresno.Text
                ps22.Fields("live").Value = TeamIndex(i)
                If TeamIndex(i) = 7 Then
                    ps22.Fields("details").Value = mydet
                End If

                ps22.Update()

            ElseIf ps22.EOF = False Then

                ' ps22.Fields("live").Value = TeamIndex(i)
                ' ps22.Fields("details").Value = mydet
                ' ps22.Update()
                'db.Execute("INSERT INTO respondent_part22(rno,live,details) VALUES('" & tresno.Text & "','" & TeamIndex(i) & "','" & mydet & "')")
                'db.Execute("UPDATE respondent_part22 SET live=" & TeamIndex(i) & ", details='" & mydet & "' where rno='" & tresno.Text & "' and reason='" & TeamIndex(i) & "'")
            End If
            ps22.Close()
        Next
    End Sub
    Sub savepart3()
        Dim ps23 As New ADODB.Recordset
        Dim mk As Integer = 0

        sopen(ps23)
        For i = 0 To dg2.Rows.Count - 1
            If dg2.Rows(i).Cells(1).Value <> Nothing Then

                If mk = 0 Then
                    db.Execute("DELETE from respondent_part3 where rno='" & tresno.Text & "'")
                    mk = 1
                End If

                ps23.Open("Select * from respondent_part3 where rno='" & tresno.Text & "' and item='" & dg2.Rows(i).Cells(0).Value & "'", db)
                If ps23.EOF = True Then

                    ps23.AddNew()
                    ps23.Fields("rno").Value = tresno.Text
                    ps23.Fields("item").Value = dg2.Rows(i).Cells(0).Value
                    ps23.Fields("rate").Value = Strings.Left(dg2.Rows(i).Cells(1).Value, 1)
                    ps23.Update()

                ElseIf ps23.EOF = False Then

                    ' ps23.Fields("rate").Value = Strings.Left(dg2.Rows(i).Cells(1).Value, 1)
                    ' ps23.Update()
                    ' db.Execute("INSERT INTO respondent_part3(rno,item,rate) VALUES('" & tresno.Text & "','" & dg2.Rows(i).Cells(0).Value & "','" & Strings.Left(dg2.Rows(i).Cells(1).Value, 1) & "')")

                    'db.Execute("UPDATE respondent_part3 SET rate=" & Strings.Left(dg2.Rows(i).Cells(1).Value, 1) & " where rno='" & tresno.Text & "' and item='" & dg2.Rows(i).Cells(0).Value & "'")
                End If
                ps23.Close()
            End If
        Next
    End Sub
    Sub savepart4()


        Dim ps4 As New ADODB.Recordset
        Dim mk As Integer = 0

        sopen(ps4)

        For j = 0 To dg3.Rows.Count - 1

            If mk = 0 Then
                db.Execute("DELETE from respondent_part41 where rno='" & tresno.Text & "'")
                mk = 1
            End If

            If dg3.Rows(j).Cells(1).Value <> Nothing Then
                ps4.Open("Select * from respondent_part41 where rno='" & tresno.Text & "' and item='" & dg3.Rows(j).Cells(0).Value & "'", db)
                If ps4.EOF = True Then
                    ps4.AddNew()
                    ps4.Fields("rno").Value = tresno.Text
                    ps4.Fields("item").Value = dg3.Rows(j).Cells(0).Value
                    ps4.Fields("rate").Value = dg3.Rows(j).Cells(1).Value
                    If dg3.Rows(j).Cells(0).Value = 9 And dg3.Rows(j).Cells(1).Value <> Nothing And dg3.Rows(j).Cells(1).Value = 1 Then
                        ps4.Fields("details").Value = t59.Text
                    End If
                    ps4.Update()
                ElseIf ps4.EOF = False Then
                    ' Dim mydet As String = Nothing

                    '  ps4.Fields("rate").Value = dg3.Rows(j).Cells(1).Value
                    ' If dg3.Rows(j).Cells(0).Value = 9 And dg3.Rows(j).Cells(1).Value <> Nothing And dg3.Rows(j).Cells(1).Value = 1 Then
                    ' ps4.Fields("details").Value = t59.Text
                    'mydet = t59.Text
                    'End If
                    ' ps4.Update()

                    ' db.Execute("INSERT INTO respondent_part41(rno,item,rate,details) VALUES('" & tresno.Text & "','" & dg3.Rows(j).Cells(0).Value & "','" & dg3.Rows(j).Cells(1).Value & "','" & mydet & "')")
                    'db.Execute("UPDATE respondent_part41 SET rate='" & dg3.Rows(j).Cells(1).Value & "', details='" & mydet & "' where rno='" & tresno.Text & "' and item='" & dg3.Rows(j).Cells(0).Value & "'")
                End If
                ps4.Close()
            End If
        Next


        mk = 0
        For j = 0 To dg4.Rows.Count - 1
            If dg4.Rows(j).Cells(1).Value <> Nothing Then

                If mk = 0 Then
                    db.Execute("DELETE from respondent_part42 where rno='" & tresno.Text & "'")
                    mk = 1
                End If

                ps4.Open("Select * from respondent_part42 where rno='" & tresno.Text & "' and item='" & dg4.Rows(j).Cells(0).Value & "'", db)
                If ps4.EOF = True Then
                    ps4.AddNew()
                    ps4.Fields("rno").Value = tresno.Text
                    ps4.Fields("item").Value = dg4.Rows(j).Cells(0).Value
                    ps4.Fields("rate").Value = dg4.Rows(j).Cells(1).Value
                    If dg4.Rows(j).Cells(0).Value = 11 And dg4.Rows(j).Cells(1).Value <> Nothing And dg4.Rows(j).Cells(1).Value = 1 Then
                        ps4.Fields("details").Value = t71.Text
                    End If
                    ps4.Update()
                ElseIf ps4.EOF = False Then
                    ' ps4.Fields("item").Value = dg4.Rows(j).Cells(0).Value
                    ' ps4.Fields("rate").Value = dg4.Rows(j).Cells(1).Value
                    '  Dim mydet As String = Nothing
                    'If dg4.Rows(j).Cells(0).Value = 11 And dg4.Rows(j).Cells(1).Value <> Nothing And dg4.Rows(j).Cells(1).Value = 1 Then
                    ' ps4.Fields("details").Value = t71.Text
                    'mydet = t71.Text
                    ' End If
                    ' ps4.Update()
                    'If mk = 0 Then
                    'db.Execute("DELETE from respondent_part42 where rno='" & tresno.Text & "'")
                    'mk = 1
                    'End If

                    'db.Execute("INSERT INTO respondent_part42(rno,item,rate,details) VALUES('" & tresno.Text & "','" & dg4.Rows(j).Cells(0).Value & "','" & dg4.Rows(j).Cells(1).Value & "','" & mydet & "')")
                    ' db.Execute("UPDATE respondent_part42 SET rate='" & dg4.Rows(j).Cells(1).Value & "', details='" & mydet & "' where rno='" & tresno.Text & "' and item='" & dg4.Rows(j).Cells(0).Value & "'")
                End If
                ps4.Close()
            End If
        Next


        Dim mydet2 As String = Nothing

        Dim TeamIndex(), i As Integer

        Dim x As Integer = 0
        ReDim Preserve TeamIndex(8)
        If t72.Checked = True Then
            TeamIndex(x) = t72.Text
            x += 1
        End If
        If t73.Checked = True Then
            TeamIndex(x) = t73.Text
            x += 1
        End If
        If t74.Checked = True Then
            TeamIndex(x) = t74.Text
            x += 1
        End If
        If t75.Checked = True Then
            TeamIndex(x) = t75.Text
            x += 1
        End If
        If t76.Checked = True Then
            TeamIndex(x) = t76.Text
            x += 1
        End If
        If t77.Checked = True Then
            TeamIndex(x) = t77.Text
            x += 1
        End If
        If t78.Checked = True Then
            TeamIndex(x) = t78.Text
            x += 1
        End If
        x = x - 1
        ReDim Preserve TeamIndex(x)

        If t78.Checked = True Then

            If t78.Text <> Nothing Then
                mydet2 = t79.Text
            Else
                mydet2 = Nothing
            End If
        End If


        mk = 0

        For i = 0 To UBound(TeamIndex)
            '   MsgBox(TeamIndex(i))
            If mk = 0 Then
                db.Execute("DELETE from respondent_part4 where rno='" & tresno.Text & "'")
                mk = 1
            End If

            ps4.Open("Select * from respondent_part4 where rno='" & tresno.Text & "' and eventno='" & TeamIndex(i) & "'", db)
            If ps4.EOF = True Then

                ps4.AddNew()
                ps4.Fields("rno").Value = tresno.Text
                ps4.Fields("eventno").Value = TeamIndex(i)
                If TeamIndex(i) = 7 Then
                    ps4.Fields("details").Value = mydet2
                End If

                ps4.Update()

            ElseIf ps4.EOF = False Then

                '  ps4.Fields("eventno").Value = TeamIndex(i)
                '  ps4.Fields("details").Value = mydet2
                ' ps4.Update()
                'If mk = 0 Then
                'db.Execute("DELETE from respondent_part4 where rno='" & tresno.Text & "'")
                'mk = 1
                'End If

                'db.Execute("INSERT INTO respondent_part4(rno,eventno,details) VALUES('" & tresno.Text & "','" & TeamIndex(i) & "','" & mydet2 & "')")
                '  db.Execute("UPDATE respondent_part4 SET eventno='" & TeamIndex(i) & "', details='" & mydet2 & "' where rno='" & tresno.Text & "' and eventno='" & TeamIndex(i) & "'")
            End If
            ps4.Close()
        Next
    End Sub

    Sub disabler()

        'dg2.Enabled = False
        'dg3.Enabled = False
        'dg4.Enabled = False


        t1.Enabled = False
        t2.Enabled = False
        t3.Enabled = False
        t4.Enabled = False
        t5.Enabled = False
        t6.Enabled = False
        t7.Enabled = False
        t8.Enabled = False
        t9.Enabled = False
        t10.Enabled = False
        t11.Enabled = False
        t12.Enabled = False
        t13.Enabled = False
        t14.Enabled = False
        t15.Enabled = False
        t16.Enabled = False
        t17.Enabled = False
        t18.Enabled = False
        t19.Enabled = False
        t20.Enabled = False


        t21.Enabled = False
        t22.Enabled = False
        t23.Enabled = False
        t24.Enabled = False
        t25.Enabled = False
        t26.Enabled = False
        t27.Enabled = False
        t28.Enabled = False
        t29.Enabled = False

        t30.Enabled = False
        t31.Enabled = False
        t32.Enabled = False
        t33.Enabled = False
        t34.Enabled = False
        t35.Enabled = False
        t36.Enabled = False
        t37.Enabled = False

        t38.Enabled = False
        t39.Enabled = False
        t40.Enabled = False
        t41.Enabled = False
        t42.Enabled = False
        t43.Enabled = False
        t44.Enabled = False
        t45.Enabled = False
        t46.Enabled = False
        t47.Enabled = False

        t72.Enabled = False
        t73.Enabled = False
        t74.Enabled = False
        t75.Enabled = False
        t76.Enabled = False
        t77.Enabled = False
        t78.Enabled = False
        t79.Enabled = False

        t48.Enabled = False
        t49.Enabled = False
        t80.Enabled = False

        For i = 0 To dg2.Rows.Count - 1
            'dg2.Rows(i).Cells(1).Value = Nothing
            dg2.CurrentCell = dg2.Rows(i).Cells(1)
            dg2.CurrentCell.ReadOnly = True
        Next
        dg2.CurrentCell = dg2.Rows(0).Cells(1)
        dg2.Focus()

        For i = 0 To dg3.Rows.Count - 1
            'dg3.Rows(i).Cells(1).Value = Nothing
            dg3.CurrentCell = dg3.Rows(i).Cells(1)
            dg3.CurrentCell.ReadOnly = True
        Next
        dg3.CurrentCell = dg3.Rows(0).Cells(1)
        dg3.Focus()
        t59.Enabled = False

        For i = 0 To dg4.Rows.Count - 1
            'dg4.Rows(i).Cells(1).Value = Nothing
            dg4.CurrentCell = dg4.Rows(i).Cells(1)
            dg4.CurrentCell.ReadOnly = True
        Next
        dg4.CurrentCell = dg4.Rows(0).Cells(1)
        dg4.Focus()
        t71.Enabled = False
    End Sub
    Sub enabler()

        dg2.Enabled = True
        dg3.Enabled = True
        dg4.Enabled = True

        t1.Enabled = True
        t2.Enabled = True
        t3.Enabled = True
        t4.Enabled = True
        t5.Enabled = True
        t6.Enabled = True
        t7.Enabled = True
        t8.Enabled = True
        t9.Enabled = True
        t10.Enabled = True
        t11.Enabled = True
        t12.Enabled = True
        t13.Enabled = True
        t14.Enabled = False
        t15.Enabled = True
        t16.Enabled = True
        t17.Enabled = True
        t18.Enabled = False
        t19.Enabled = True
        t20.Enabled = False
        t20.Text = Nothing

        t21.Enabled = True
        t22.Enabled = True
        t23.Enabled = True
        t24.Enabled = True
        t25.Enabled = True
        t26.Enabled = True
        t27.Enabled = True
        t28.Enabled = True
        t29.Enabled = False
        t29.Text = Nothing

        t30.Enabled = True
        t31.Enabled = True
        t32.Enabled = True
        t33.Enabled = True
        t34.Enabled = True
        t35.Enabled = True
        t36.Enabled = True
        t37.Enabled = False

        t38.Enabled = True
        t39.Enabled = True
        t40.Enabled = True
        t41.Enabled = True
        t42.Enabled = True
        t43.Enabled = True

        t44.Enabled = True
        t45.Enabled = True
        t46.Enabled = True
        t47.Enabled = True
        t59.Text = Nothing
        t59.Enabled = False

        t71.Text = Nothing
        t71.Enabled = False

        t72.Enabled = True
        t73.Enabled = True
        t74.Enabled = True
        t75.Enabled = True
        t76.Enabled = True
        t77.Enabled = True
        t78.Enabled = True
        t79.Text = Nothing
        t79.Enabled = False

        t48.Enabled = True
        t49.Enabled = True
        t80.Enabled = True

        For i = 0 To dg2.Rows.Count - 1
            dg2.Rows(i).Cells(1).Value = Nothing
        Next
        dg2.CurrentCell = dg2.Rows(0).Cells(1)
        dg2.CurrentCell.ReadOnly = False
        dg2.Focus()

        For i = 0 To dg3.Rows.Count - 1
            dg3.Rows(i).Cells(1).Value = Nothing
        Next
        dg3.CurrentCell = dg3.Rows(0).Cells(1)
        dg3.CurrentCell.ReadOnly = False
        dg3.Focus()

        For i = 0 To dg4.Rows.Count - 1
            dg4.Rows(i).Cells(1).Value = Nothing
        Next
        dg4.CurrentCell = dg4.Rows(0).Cells(1)
        dg4.CurrentCell.ReadOnly = False
        dg4.Focus()

        lblrno.Text = Nothing
        lblrn.Visible = False
        lblrno.Visible = False
        tcr.SelectedIndex = 0
    End Sub
    Sub clearer()

        t1.Text = Nothing
        t2.Text = Nothing
        t3.Text = Nothing
        t4.Text = Nothing
        t5.Text = Nothing
        t6.Text = Nothing
        t7.Text = Nothing
        t8.Text = Nothing
        t9.Text = Nothing
        t10.Text = Nothing
        t11.Text = Nothing
        t12.Text = Nothing
        t13.Text = Nothing
        t14.Text = Nothing
        t15.Text = Nothing
        t16.Text = Nothing
        t17.Text = Nothing
        t18.Text = Nothing
        t19.Text = Nothing
        t20.Text = Nothing


        t21.Checked = False
        t22.Checked = False
        t23.Checked = False
        t24.Checked = False
        t25.Checked = False
        t26.Checked = False
        t27.Checked = False
        t28.Checked = False
        t29.Text = Nothing

        t30.Checked = False
        t31.Checked = False
        t32.Checked = False
        t33.Checked = False
        t34.Checked = False
        t35.Checked = False
        t36.Checked = False
        t37.Text = Nothing

        t38.Checked = False
        t39.Checked = False
        t40.Checked = False
        t41.Checked = False
        t42.Checked = False
        t43.Checked = False

        t44.Checked = False
        t45.Checked = False
        t46.Checked = False
        t47.Checked = False
        t59.Text = Nothing
        t59.Enabled = False

        t71.Text = Nothing
        t71.Enabled = False

        t72.Checked = False
        t73.Checked = False
        t74.Checked = False
        t75.Checked = False
        t76.Checked = False
        t77.Checked = False
        t78.Checked = False
        t79.Text = Nothing
        t79.Enabled = False

        t48.Text = Nothing
        t49.Text = Nothing
        t80.Text = Nothing

    End Sub

    Private Sub resetme_Click(sender As Object, e As EventArgs)
        clearer()
        getautoincrement()
    End Sub
    Sub getautoincrement()
        Dim gat As New ADODB.Recordset
        sopen(gat)
        gat.Open("SHOW TABLE STATUS WHERE `Name` = 'respondent'", db)
        tresno.Text = gat.Fields("Auto_increment").Value
        gat.Close()
    End Sub


    Private Sub t28_CheckedChanged(sender As Object, e As EventArgs) Handles t28.CheckedChanged
        If t28.Checked = True Then
            t29.Enabled = True
        Else
            t29.Enabled = False
            t29.Text = Nothing
        End If
    End Sub
    Private Sub t36_CheckedChanged(sender As Object, e As EventArgs) Handles t36.CheckedChanged
        If t36.Checked = True Then
            t37.Enabled = True
        Else
            t37.Enabled = False
            t37.Text = Nothing
        End If
    End Sub

    Private Sub t13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles t13.SelectedIndexChanged
        If t13.Text = "No" Then
            t14.Enabled = True
        ElseIf t13.Text = "Yes" Then
            t14.Enabled = False
            t14.Text = Nothing
        Else
            t14.Enabled = False
            t14.Text = Nothing
        End If
    End Sub

    Private Sub t17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles t17.SelectedIndexChanged
        If t17.Text = "Yes" Then
            t18.Enabled = True
        ElseIf t17.Text = "No" Then
            t18.Enabled = False
            t18.Text = Nothing
        Else
            t18.Enabled = False
            t18.Text = Nothing
        End If
    End Sub

    Private Sub t19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles t19.SelectedIndexChanged
        If t19.Text = "Yes" Then
            t20.Enabled = True
        ElseIf t19.Text = "No" Then
            t20.Enabled = False
            t20.Text = Nothing
        Else
            t20.Enabled = False
            t20.Text = Nothing
        End If
    End Sub

    Private Sub t78_CheckedChanged(sender As Object, e As EventArgs) Handles t78.CheckedChanged
        If t78.Checked = True Then
            t79.Enabled = True
        Else
            t79.Enabled = False
            t79.Text = Nothing
        End If
    End Sub

    Private Sub dg3_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dg3.CellEndEdit

    End Sub


    Private Sub dg3_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg3.CellValueChanged
        If dg3.Rows.Count > 0 Then
            If dg3.Rows(dg3.Rows.Count - 1).Cells(0).Value = 9 And dg3.Rows(dg3.Rows.Count - 1).Cells(1).Value <> Nothing And dg3.Rows(dg3.Rows.Count - 1).Cells(1).Value = 1 Then
                t59.Enabled = True
            Else
                t59.Text = Nothing
                t59.Enabled = False
            End If
        End If

    End Sub

    Private Sub dg4_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dg4.CellValueChanged
        If dg4.Rows.Count > 0 Then
            If dg4.Rows(dg4.Rows.Count - 1).Cells(0).Value = 11 And dg4.Rows(dg4.Rows.Count - 1).Cells(1).Value <> Nothing And dg4.Rows(dg4.Rows.Count - 1).Cells(1).Value = 1 Then
                t71.Enabled = True
            Else
                t71.Text = Nothing
                t71.Enabled = False
            End If
        End If
    End Sub


    Private Sub tsreport_Click(sender As Object, e As EventArgs) Handles tsreport.Click

        svid = 4
        With frmfrsreports
            .MdiParent = DSEGMAIN
            ' .Dock = DockStyle.Fill
            .BringToFront()
            .Show()
        End With

    End Sub


    Private Sub tcr_KeyDown(sender As Object, e As KeyEventArgs) Handles tcr.KeyDown
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then
            saveme.PerformClick()
        End If
    End Sub



    Private Sub dg2_KeyDown(sender As Object, e As KeyEventArgs) Handles dg2.KeyDown
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then
            saveme.PerformClick()
        End If
    End Sub


    Private Sub dg3_KeyDown(sender As Object, e As KeyEventArgs) Handles dg3.KeyDown
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then
            saveme.PerformClick()
        End If
    End Sub

    Private Sub dg4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg4.CellContentClick

    End Sub

    Private Sub dg4_KeyDown(sender As Object, e As KeyEventArgs) Handles dg4.KeyDown
        If (e.KeyCode = Keys.S AndAlso e.Modifiers = Keys.Control) Then
            saveme.PerformClick()
        End If
    End Sub

    Private Sub tsrespondent_Click(sender As Object, e As EventArgs) Handles tsrespondent.Click
        With frmrespondents
            .lblsno.Text = lblsurveyid.Text
            .Show(Me)
        End With
    End Sub


    Private Sub tsreset_Click(sender As Object, e As EventArgs) Handles tsreset.Click
        clearer()
        enabler()
        t1.Focus()
    End Sub



End Class