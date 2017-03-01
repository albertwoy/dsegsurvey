Public Class frmLogin

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim uchk As New ADODB.Recordset
        sopen(uchk)
        uchk.Open("Select * from eguidance.login_info where BINARY(username)='" & cq(tuser.Text) & "' and BINARY(passwd)='" & cq(tpass.Text) & "'", db)
        If uchk.EOF = False Then
            If uchk.Fields("active").Value = "1" Then
                MsgBox("Access Granted!", 64, "THANK YOU!")
                With DSEGMAIN
                    .menuLout.Enabled = True
                    .menuFile.Enabled = True
                    .menuManage.Enabled = True
                    .menuLin.Enabled = False
                    .tsstatus.Text = tuser.Text
                End With
                Me.Close()
            Else
                MsgBox("This user account is inactive, please use an active account", 48, "Account inactive")
            End If
        Else
            MsgBox("Unknown user information", 48, "UNKNOWN USER")
        End If
        uchk.Close()


    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
