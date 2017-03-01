Imports System.ComponentModel
Imports System.Threading
Public Class frmrespondents

   Private Delegate Function ReturnDelegate() As Object
    Private row As New DataGridViewRow
    Dim r1 As System.Threading.Thread
    Dim selrno As String = Nothing
    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    Sub disp()
        Dim rn As New ADODB.Recordset
        sopen(rn)
        rn.Open("Select * from respondent where surveyid='" & lblsno.Text & "' ORDER BY rno DESC LIMIT 1000", db)
        While Not rn.EOF
            row = New DataGridViewRow()
            row.CreateCells(dgres)
            row.Cells(0).Value = rn.Fields("rno").Value
            row.Cells(1).Value = rn.Fields("rname").Value
            row.Cells(2).Value = rn.Fields("course").Value
            row.Cells(3).Value = rn.Fields("year").Value
            row.Cells(4).Value = rn.Fields("gender").Value
            row.Cells(5).Value = rn.Fields("age").Value
            row.Cells(6).Value = rn.Fields("relg").Value
            row.Cells(7).Value = rn.Fields("abused").Value
            row.Cells(8).Value = rn.Fields("abusedy").Value
            row.Cells(9).Value = rn.Fields("msem").Value & "-" & rn.Fields("myear").Value
            row.Cells(10).Value = "select"
            AddRow()
            rn.MoveNext()
        End While
        rn.Close()
        Me.Cursor = Cursors.Hand

    End Sub
    Private Overloads Function AddRow() As Integer

        If InvokeRequired Then
            Return CInt(Invoke(New ReturnDelegate(AddressOf AddRow)))
        Else
            Return dgres.Rows.Add(row)
        End If

    End Function

    Private Sub frmrespondents_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If Not r1 Is Nothing Then
            r1.Abort()
        End If
    End Sub

    Private Sub frmrespondents_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        Me.Cursor = Cursors.AppStarting
        r1 = Nothing
        If r1 Is Nothing Then
            r1 = New Thread(Sub() disp())
            r1.IsBackground = True
            r1.Start()
        End If
    End Sub

    Private Sub dgres_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgres.CellContentClick
        If e.RowIndex > -1 And dgres.Rows.Count > 0 And e.ColumnIndex = 10 Then
            'r1 = Nothing
            'If r1 Is Nothing Then
            selrno = dgres.Rows(e.RowIndex).Cells(0).Value
            '    r1 = New Thread(Sub() retriever())
            '    r1.IsBackground = True
            '    r1.Start()
            'End If
            Me.Cursor = Cursors.AppStarting
            retriever()
        End If
    End Sub

    Sub retriever()
        If svid = 4 Then
            rtr2()
        Else
            rtr()

        End If
        Me.Cursor = Cursors.Hand


    End Sub
    Sub rtr2()

        With frmfrs
            .clearer()
            .enabler()
            .lblrn.Visible = True
            .lblrno.Visible = True
            .lblrno.Text = selrno
            Dim rf As New ADODB.Recordset
            sopen(rf)
            rf.Open("Select * from respondent where surveyid='" & svid & "' and rno='" & selrno & "'", db)
            If rf.EOF = False Then

                .t1.Text = rf.Fields("rname").Value

                Dim x As Integer = 0
                Dim mstr As String = Nothing

                mstr = IIf(IsDBNull(rf.Fields("course").Value), Nothing, rf.Fields("course").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t2.FindString(mstr)
                    .t2.SelectedIndex = x
                End If


                x = 0
                mstr = IIf(IsDBNull(rf.Fields("year").Value), Nothing, rf.Fields("year").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t3.FindString(mstr)
                    .t3.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("gender").Value), Nothing, rf.Fields("gender").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t4.FindString(mstr)
                    .t4.SelectedIndex = x
                End If

                .t5.Text = IIf(IsDBNull(rf.Fields("age").Value), Nothing, rf.Fields("age").Value)
                .t6.Text = IIf(IsDBNull(rf.Fields("wgt").Value), Nothing, rf.Fields("wgt").Value)
                .t7.Text = IIf(IsDBNull(rf.Fields("hgt").Value), Nothing, rf.Fields("hgt").Value)

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("relg").Value), Nothing, rf.Fields("relg").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t8.FindString(mstr)
                    .t8.SelectedIndex = x
                End If

                .t9.Text = IIf(IsDBNull(rf.Fields("totsib").Value), Nothing, rf.Fields("totsib").Value)
                .t10.Text = IIf(IsDBNull(rf.Fields("totfam").Value), Nothing, rf.Fields("totfam").Value)

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("youare").Value), Nothing, rf.Fields("youare").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t11.FindString(mstr)
                    .t11.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("youpar").Value), Nothing, rf.Fields("youpar").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t12.FindString(mstr)
                    .t12.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("parlive").Value), Nothing, rf.Fields("parlive").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t13.FindString(mstr)
                    .t13.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("deceased").Value), Nothing, rf.Fields("deceased").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t14.FindString(mstr)
                    .t14.SelectedIndex = x
                End If


                x = 0
                mstr = IIf(IsDBNull(rf.Fields("livin").Value), Nothing, rf.Fields("livin").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t15.FindString(mstr)
                    .t15.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("insrc").Value), Nothing, rf.Fields("insrc").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t16.FindString(mstr)
                    .t16.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("abused").Value), Nothing, rf.Fields("abused").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t17.FindString(mstr)
                    .t17.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("abusedy").Value), Nothing, rf.Fields("abusedy").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t18.FindString(mstr)
                    .t18.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rf.Fields("negl").Value), Nothing, rf.Fields("negl").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .t19.FindString(mstr)
                    .t19.SelectedIndex = x
                End If

                .t20.Text = IIf(IsDBNull(rf.Fields("negly").Value), Nothing, rf.Fields("negly").Value)

                Dim metm As String = Nothing
                Dim stm As String = Nothing

                stm = IIf(IsDBNull(rf.Fields("studmu").Value), Nothing, rf.Fields("studmu").Value)
                metm = IIf(IsDBNull(rf.Fields("lastsem").Value), Nothing, rf.Fields("lastsem").Value)

                If stm = .t38.Text Then
                    .t38.Checked = True
                ElseIf stm = .t39.Text Then
                    .t39.Checked = True
                ElseIf stm = .t40.Text Then
                    .t40.Checked = True
                ElseIf stm = .t41.Text Then
                    .t41.Checked = True
                ElseIf stm = .t42.Text Then
                    .t42.Checked = True
                ElseIf stm = .t43.Text Then
                    .t43.Checked = True
                End If

                If metm = .t44.Text Then
                    .t44.Checked = True
                ElseIf metm = .t45.Text Then
                    .t45.Checked = True
                ElseIf metm = .t46.Text Then
                    .t46.Checked = True
                ElseIf metm = .t47.Text Then
                    .t47.Checked = True
                End If

                .t48.Text = IIf(IsDBNull(rf.Fields("nummet").Value), Nothing, rf.Fields("nummet").Value)
                .t49.Text = IIf(IsDBNull(rf.Fields("numclass").Value), Nothing, rf.Fields("numclass").Value)
                .t80.Text = IIf(IsDBNull(rf.Fields("anything").Value), Nothing, rf.Fields("anything").Value)

            End If
            rf.Close()

            .tcr.SelectedIndex = 1
            mytab = 1
            rf.Open("Select * from respondent_part21  where rno='" & selrno & "'", db)
            While Not rf.EOF
                If IsDBNull(rf.Fields("reason").Value) Then
                Else
                    Dim mres As Integer = 0
                    mres = rf.Fields("reason").Value
                    If mres = .t21.Text Then
                        .t21.Checked = True
                    ElseIf mres = .t22.Text Then
                        .t22.Checked = True
                    ElseIf mres = .t23.Text Then
                        .t23.Checked = True
                    ElseIf mres = .t24.Text Then
                        .t24.Checked = True
                    ElseIf mres = .t25.Text Then
                        .t25.Checked = True
                    ElseIf mres = .t26.Text Then
                        .t26.Checked = True
                    ElseIf mres = .t27.Text Then
                        .t27.Checked = True
                    ElseIf mres = .t28.Text Then
                        .t28.Checked = True
                        .t29.Text = IIf(IsDBNull(rf.Fields("details").Value), Nothing, rf.Fields("details").Value)
                        '  .t29.Enabled = False
                    End If

                    If .t28.Checked = True Then
                        .t29.Enabled = True
                    Else
                        .t29.Enabled = False
                    End If
                End If
                rf.MoveNext()
            End While
            rf.Close()

            rf.Open("Select * from respondent_part22 where rno='" & selrno & "'", db)
            While Not rf.EOF
                If IsDBNull(rf.Fields("live").Value) Then
                Else
                    Dim mres As Integer = 0
                    mres = rf.Fields("live").Value
                    If mres = .t30.Text Then
                        .t30.Checked = True
                    ElseIf mres = .t31.Text Then
                        .t31.Checked = True
                    ElseIf mres = .t32.Text Then
                        .t32.Checked = True
                    ElseIf mres = .t33.Text Then
                        .t33.Checked = True
                    ElseIf mres = .t34.Text Then
                        .t34.Checked = True
                    ElseIf mres = .t35.Text Then
                        .t35.Checked = True
                    ElseIf mres = .t36.Text Then
                        .t36.Checked = True
                        .t37.Text = IIf(IsDBNull(rf.Fields("details").Value), Nothing, rf.Fields("details").Value)
                        '.t37.Enabled = False
                    End If
                    If .t36.Checked = True Then
                        .t37.Enabled = True
                    Else
                        .t37.Enabled = False
                    End If
                End If
                rf.MoveNext()
            End While
            rf.Close()

            .tcr.SelectedIndex = 2
            mytab = 2
            .dg2.Focus()
            For i = 0 To .dg2.Rows.Count - 1
                Dim qid As String = Nothing
                qid = .dg2.Rows(i).Cells(0).Value

                rf.Open("Select * from respondent_part3 where rno='" & selrno & "' and item='" & qid & "'", db)
                If rf.EOF = False Then
                    If IsDBNull(rf.Fields("rate").Value) Then
                        .dg2.Rows(i).Cells(1).Value = Nothing
                        ' .dgq1.Rows(i).Cells(1).Value = Strings.Left(rd.Fields("rate").Value, 1)
                    Else

                        .dg2.CurrentCell = .dg2.Rows(i).Cells(1)
                        .dg2.CurrentCell.ReadOnly = False
                        .dg2.Focus()
                        .dg2.BeginEdit(True)
                        ' Thread.Sleep(100)
                        SendKeys.SendWait(Strings.Left(rf.Fields("rate").Value, 1))

                    End If
                Else
                    .dg2.Rows(i).Cells(1).Value = Nothing
                End If
                rf.Close()

            Next


            .tcr.SelectedIndex = 3
            mytab = 3
            rf.Open("Select * from respondent_part4 where rno='" & selrno & "'", db)
            While Not rf.EOF
                If IsDBNull(rf.Fields("eventno").Value) Then
                Else
                    Dim mres As Integer = 0
                    mres = rf.Fields("eventno").Value
                    If mres = .t72.Text Then
                        .t72.Checked = True
                    ElseIf mres = .t73.Text Then
                        .t73.Checked = True
                    ElseIf mres = .t74.Text Then
                        .t74.Checked = True
                    ElseIf mres = .t75.Text Then
                        .t75.Checked = True
                    ElseIf mres = .t76.Text Then
                        .t76.Checked = True
                    ElseIf mres = .t77.Text Then
                        .t77.Checked = True
                    ElseIf mres = .t78.Text Then
                        .t78.Checked = True
                        .t79.Text = IIf(IsDBNull(rf.Fields("details").Value), Nothing, rf.Fields("details").Value)
                        ' .t79.Enabled = False
                    End If
                    If .t78.Checked = True Then
                        .t79.Enabled = True
                    Else
                        .t79.Enabled = False
                    End If
                End If
                rf.MoveNext()
            End While
            rf.Close()

            .dg3.Focus()
            For i = 0 To .dg3.Rows.Count - 1
                Dim qid As String = Nothing
                qid = .dg3.Rows(i).Cells(0).Value

                rf.Open("Select * from respondent_part41 where rno='" & selrno & "' and item='" & qid & "'", db)
                If rf.EOF = False Then
                    If IsDBNull(rf.Fields("rate").Value) Then
                        .dg3.Rows(i).Cells(1).Value = Nothing
                        ' .dgq1.Rows(i).Cells(1).Value = Strings.Left(rd.Fields("rate").Value, 1)
                    Else


                        .dg3.CurrentCell = .dg3.Rows(i).Cells(1)
                        .dg3.CurrentCell.ReadOnly = False
                        .dg3.Focus()
                        .dg3.BeginEdit(True)
                        ' Thread.Sleep(100)
                        SendKeys.SendWait(Strings.Left(rf.Fields("rate").Value, 1))

                        If .dg3.Rows(i).Cells(0).Value = 9 And .dg3.Rows(i).Cells(1).Value <> Nothing And .dg3.Rows(i).Cells(1).Value = 1 Then
                            .t59.Text = IIf(IsDBNull(rf.Fields("details").Value), Nothing, rf.Fields("details").Value)
                        End If

                    End If
                Else
                    .dg3.Rows(i).Cells(1).Value = Nothing
                End If
                rf.Close()

            Next


            .tcr.SelectedIndex = 3
            mytab = 3
            .dg4.Focus()
            For i = 0 To .dg4.Rows.Count - 1
                Dim qid As String = Nothing
                qid = .dg4.Rows(i).Cells(0).Value

                rf.Open("Select * from respondent_part42 where rno='" & selrno & "' and item='" & qid & "'", db)
                If rf.EOF = False Then
                    If IsDBNull(rf.Fields("rate").Value) Then
                        .dg4.Rows(i).Cells(1).Value = Nothing
                    Else

                        .dg4.CurrentCell = .dg4.Rows(i).Cells(1)
                        .dg4.CurrentCell.ReadOnly = False
                        .dg4.Focus()
                        .dg4.BeginEdit(True)
                        ' Thread.Sleep(100)
                        SendKeys.SendWait(Strings.Left(rf.Fields("rate").Value, 1))

                        If .dg4.Rows(i).Cells(0).Value = 11 And .dg4.Rows(i).Cells(1).Value <> Nothing And .dg4.Rows(i).Cells(1).Value = 1 Then
                            .t71.Text = IIf(IsDBNull(rf.Fields("details").Value), Nothing, rf.Fields("details").Value)
                        End If

                    End If
                Else
                    .dg4.Rows(i).Cells(1).Value = Nothing
                End If
                rf.Close()

            Next
            Me.Close()
            '.disabler()
            .tcr.SelectedIndex = 0
            mytab = 0

        End With

    End Sub

    Sub rtr()

        With frmsurveyform
            .cbocrs1.Text = Nothing
            .cboyr1.Text = Nothing
            .cbogender.Text = Nothing
            .dgq1.Text = Nothing
            .txtage.Text = Nothing
            .txthome.Text = Nothing

            .lblrn.Visible = True
            .lblrno.Visible = True
            .lblrno.Text = selrno

            Dim rd As New ADODB.Recordset
            sopen(rd)
            rd.Open("Select * from respondent where surveyid='" & svid & "' and rno='" & selrno & "'", db)
            If rd.EOF = False Then

                Dim x As Integer = 0
                Dim mstr As String = Nothing
                mstr = IIf(IsDBNull(rd.Fields("course").Value), Nothing, rd.Fields("course").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .cbocrs1.FindString(mstr)
                    .cbocrs1.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rd.Fields("year").Value), Nothing, rd.Fields("year").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .cboyr1.FindString(mstr)
                    .cboyr1.SelectedIndex = x
                End If

                x = 0
                mstr = IIf(IsDBNull(rd.Fields("gender").Value), Nothing, rd.Fields("gender").Value)
                If mstr <> Nothing And mstr <> "" Then
                    x = .cbogender.FindString(mstr)
                    .cbogender.SelectedIndex = x
                End If

                .txtage.Text = IIf(IsDBNull(rd.Fields("age").Value), Nothing, rd.Fields("age").Value)
                .txthome.Text = IIf(IsDBNull(rd.Fields("homeaddress").Value), Nothing, rd.Fields("homeaddress").Value)

            End If
            rd.Close()

            For i = 0 To .dgq1.Rows.Count - 1
                Dim cid As String = Nothing
                Dim qid As String = Nothing

                cid = .dgq1.Rows(i).Cells(2).Value
                qid = .dgq1.Rows(i).Cells(0).Value

                rd.Open("Select * from tab_surveyans where surveyid='" & svid & "' and rno='" & selrno & "' and catid='" & cid & "' and qid=" & qid & "", db)
                If rd.EOF = False Then
                    If IsDBNull(rd.Fields("rate").Value) Then
                        .dgq1.Rows(i).Cells(1).Value = Nothing
                        ' .dgq1.Rows(i).Cells(1).Value = Strings.Left(rd.Fields("rate").Value, 1)
                    Else
                        Thread.Sleep(30)
                        .dgq1.CurrentCell = .dgq1.Rows(i).Cells(1)
                        .dgq1.CurrentCell.ReadOnly = False
                        .dgq1.Focus()
                        .dgq1.BeginEdit(True)
                        SendKeys.SendWait(Strings.Left(rd.Fields("rate").Value, 1))

                    End If
                Else
                    .dgq1.Rows(i).Cells(1).Value = Nothing
                End If
                rd.Close()
                Me.Close()
            Next

            '.cbocrs1.Enabled = False
            '.cboyr1.Enabled = False
            '.cbogender.Enabled = False
            '.dgq1.Enabled = False
            '.txtage.Enabled = False
            '.txthome.Enabled = False

        End With

    End Sub
   
    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If txtsearch.Text <> Nothing Then
            Dim qstr As String = Nothing
            qstr = cq(txtsearch.Text)
            dgres.Rows.Clear()
            Me.Cursor = Cursors.AppStarting
            Dim rp As New ADODB.Recordset
            sopen(rp)
            rp.Open("Select * from respondent where CONCAT_WS(' ', rname, course, year, gender, age, relg) LIKE '%" & qstr & "%' and surveyid='" & lblsno.Text & "'  ORDER BY rno DESC LIMIT 100", db)
            While Not rp.EOF
                row = New DataGridViewRow()
                row.CreateCells(dgres)
                row.Cells(0).Value = rp.Fields("rno").Value
                row.Cells(1).Value = rp.Fields("rname").Value
                row.Cells(2).Value = rp.Fields("course").Value
                row.Cells(3).Value = rp.Fields("year").Value
                row.Cells(4).Value = rp.Fields("gender").Value
                row.Cells(5).Value = rp.Fields("age").Value
                row.Cells(6).Value = rp.Fields("relg").Value
                row.Cells(7).Value = rp.Fields("abused").Value
                row.Cells(8).Value = rp.Fields("abusedy").Value
                row.Cells(9).Value = rp.Fields("msem").Value & "-" & rp.Fields("myear").Value
                row.Cells(10).Value = "select"
                AddRow()
                rp.MoveNext()
            End While
            rp.Close()
            Me.Cursor = Cursors.Hand

        Else
            Me.Cursor = Cursors.AppStarting
            r1 = Nothing
            If r1 Is Nothing Then
                r1 = New Thread(Sub() disp())
                r1.IsBackground = True
                r1.Start()
            End If

        End If
    End Sub
End Class