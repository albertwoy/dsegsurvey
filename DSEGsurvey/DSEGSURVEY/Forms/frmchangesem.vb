Public Class frmchangesem
    Dim ms As New ADODB.Recordset
    Private Sub dgsem_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgsem.CellContentClick
        If dgsem.Rows.Count > -1 Then
            If e.ColumnIndex = 3 Then
                With DSEGMAIN
                    .tssem.Text = dgsem.Rows(e.RowIndex).Cells(1).Value
                    .tsyear.Text = dgsem.Rows(e.RowIndex).Cells(2).Value
                End With

                Me.Close()
            ElseIf e.ColumnIndex = 4 Then
                Dim sid As Integer = 0
                sid = dgsem.Rows(e.RowIndex).Cells(0).Value
                Dim c As Integer = 0
                If MsgBox("Are you sure you want to delete|remove this semester?", 4 + 32, "REMOVE SELECTED SEMESTER?") = 7 Then
                    c = 1
                Else
                    db.Execute("Delete from sched where ID=" & sid & "")
                    displaysems()
                End If
            End If
        End If
    End Sub

    Private Sub frmchangesem_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Control AndAlso e.KeyCode = 83 Then
            saveme.PerformClick()
        End If
    End Sub

    Private Sub frmchangesem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cboyear.Items.Clear()
        Dim myyr As Integer = 0
        myyr = gcyear()
        myyr = myyr + 1
        For i = 2010 To myyr
            cboyear.Items.Add(i)
        Next
        displaysems()

    End Sub
    Sub displaysems()
        dgsem.Rows.Clear()
        Dim row As New DataGridViewRow

        sopen(ms)
        ms.Open("Select * from sched order by ID DESC", db)
        While Not ms.EOF
            row = New DataGridViewRow()
            row.CreateCells(dgsem)
            row.Cells(0).Value = ms.Fields("ID").Value
            row.Cells(1).Value = ms.Fields("msem").Value
            row.Cells(2).Value = ms.Fields("myear").Value
            row.Cells(3).Value = "Select"
            row.Cells(4).Value = "Remove"
            dgsem.Rows.Add(row)
            ms.MoveNext()
        End While
       

        ms.Close()
    End Sub

    Private Sub saveme_Click(sender As Object, e As EventArgs) Handles saveme.Click
        If cbosem.Text <> Nothing And cboyear.Text <> Nothing Then
            savesems()
        Else
            MsgBox("Please select and semester first", 48, "SELECT FIRST")
        End If

    End Sub
    Sub savesems()
        sopen(ms)
        ms.Open("Select * from sched", db)
        ms.AddNew()
        ms.Fields("msem").Value = cbosem.Text
        ms.Fields("myear").Value = cboyear.Text
        ms.Update()
        ms.Close()
        displaysems()
    End Sub

    Private Sub tsrefresh_Click(sender As Object, e As EventArgs) Handles tsrefresh.Click
        displaysems()
    End Sub
End Class