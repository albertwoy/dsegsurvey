Imports System.Windows.Forms

Public Class DSEGMAIN

    Private Sub menuNA_Click(sender As Object, e As EventArgs) Handles menuNA.Click
        closeallchild()
        svid = 1
        With frmsurveyform
            .lblsurveyid.Text = svid
            .Text = .Text & " [" & Strings.Mid(menuNA.Text, 2) & "]"
            cmenu = Strings.Mid(menuNA.Text, 2)
            .MdiParent = Me
            ' .Dock = DockStyle.Fill
            .Show()
        End With

    End Sub

    Private Sub menuCES_Click(sender As Object, e As EventArgs) Handles menuCES.Click
        closeallchild()
        svid = 2
        With frmsurveyform
            .lblsurveyid.Text = svid
            .Text = .Text & " [" & Strings.Mid(menuCES.Text, 2) & "]"
            cmenu = Strings.Mid(menuCES.Text, 2)
            .MdiParent = Me
            ' .Dock = DockStyle.Fill
            .Show()
        End With
    End Sub

    Private Sub menuGS_Click(sender As Object, e As EventArgs) Handles menuGS.Click
        closeallchild()
        svid = 3
        With frmsurveyform
            .lblsurveyid.Text = svid
            .Text = .Text & " [" & Strings.Mid(menuGS.Text, 2) & "]"
            cmenu = Strings.Mid(menuGS.Text, 2)
            .MdiParent = Me
            ' .Dock = DockStyle.Fill
            .Show()
        End With

    End Sub

    Private Sub menuFRS_Click(sender As Object, e As EventArgs) Handles menuFRS.Click
        closeallchild()
        svid = 4
        With frmfrs
            .lblsurveyid.Text = svid
            .Text = .Text & " [" & Strings.Mid(menuFRS.Text, 2) & "]"
            cmenu = Strings.Mid(menuFRS.Text, 2)
            .MdiParent = Me
            ' .Dock = DockStyle.Fill
            .Show()
        End With

    End Sub

    Private Sub DSEGMAIN_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connecttodb()
        Dim csem As New ADODB.Recordset
        sopen(csem)
        csem.Open("Select * from sched order by ID DESC", db)
        tssem.Text = csem.Fields("msem").Value
        tsyear.Text = csem.Fields("myear").Value
        csem.Close()
    End Sub

 
    Private Sub menuManagesem_Click(sender As Object, e As EventArgs) Handles menuManagesem.Click
        closeallchild()
        With frmchangesem
            .Text = .Text & " [" & Strings.Mid(menuManagesem.Text, 2) & "]"
            .MdiParent = Me
            ' .Dock = DockStyle.Fill
            .Show()
        End With
    End Sub

    Private Sub menuLout_Click(sender As Object, e As EventArgs) Handles menuLout.Click
        tsstatus.Text = "user account"
        menuLout.Enabled = False
        menuFile.Enabled = False
        menuManage.Enabled = False
        menuLin.Enabled = True
        frmLogin.Show(Me)
    End Sub

    Private Sub menuLin_Click(sender As Object, e As EventArgs) Handles menuLin.Click
      
        frmLogin.Show(Me)
    End Sub
End Class
