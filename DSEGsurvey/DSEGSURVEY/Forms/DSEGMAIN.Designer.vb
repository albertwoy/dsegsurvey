<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DSEGMAIN
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub


    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DSEGMAIN))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.menuUser = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuLout = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuLin = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuNS = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuNA = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuCES = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuGS = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuFRS = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuManage = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuManagesem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.tsstatus = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tssem = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsyear = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tstat1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuUser, Me.menuFile, Me.menuManage})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Size = New System.Drawing.Size(881, 24)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'menuUser
        '
        Me.menuUser.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuLout, Me.menuLin})
        Me.menuUser.Name = "menuUser"
        Me.menuUser.Size = New System.Drawing.Size(42, 20)
        Me.menuUser.Text = "Use&r"
        '
        'menuLout
        '
        Me.menuLout.Enabled = False
        Me.menuLout.Name = "menuLout"
        Me.menuLout.Size = New System.Drawing.Size(152, 22)
        Me.menuLout.Text = "Log&out"
        '
        'menuLin
        '
        Me.menuLin.Name = "menuLin"
        Me.menuLin.Size = New System.Drawing.Size(152, 22)
        Me.menuLin.Text = "Log&in"
        '
        'menuFile
        '
        Me.menuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuNS})
        Me.menuFile.Enabled = False
        Me.menuFile.Name = "menuFile"
        Me.menuFile.Size = New System.Drawing.Size(37, 20)
        Me.menuFile.Text = "&File"
        '
        'menuNS
        '
        Me.menuNS.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuNA, Me.menuCES, Me.menuGS, Me.menuFRS})
        Me.menuNS.Name = "menuNS"
        Me.menuNS.Size = New System.Drawing.Size(152, 22)
        Me.menuNS.Text = "&New Survey"
        '
        'menuNA
        '
        Me.menuNA.Name = "menuNA"
        Me.menuNA.Size = New System.Drawing.Size(227, 22)
        Me.menuNA.Text = "&Needs Assessment"
        '
        'menuCES
        '
        Me.menuCES.Name = "menuCES"
        Me.menuCES.Size = New System.Drawing.Size(227, 22)
        Me.menuCES.Text = "&Campus Environment Survey"
        '
        'menuGS
        '
        Me.menuGS.Name = "menuGS"
        Me.menuGS.Size = New System.Drawing.Size(227, 22)
        Me.menuGS.Text = "&Guidance Services"
        '
        'menuFRS
        '
        Me.menuFRS.Name = "menuFRS"
        Me.menuFRS.Size = New System.Drawing.Size(227, 22)
        Me.menuFRS.Text = "&Freshmen Retention Survey"
        '
        'menuManage
        '
        Me.menuManage.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuManagesem})
        Me.menuManage.Enabled = False
        Me.menuManage.Name = "menuManage"
        Me.menuManage.Size = New System.Drawing.Size(62, 20)
        Me.menuManage.Text = "&Manage"
        '
        'menuManagesem
        '
        Me.menuManagesem.Name = "menuManagesem"
        Me.menuManagesem.Size = New System.Drawing.Size(166, 22)
        Me.menuManagesem.Text = "&Change Semester"
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsstatus, Me.ToolStripStatusLabel2, Me.tssem, Me.ToolStripStatusLabel1, Me.tsyear})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 431)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(881, 22)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'tsstatus
        '
        Me.tsstatus.Name = "tsstatus"
        Me.tsstatus.Size = New System.Drawing.Size(118, 17)
        Me.tsstatus.Text = "Current user account"
        Me.tsstatus.ToolTipText = "current user account"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(196, 17)
        Me.ToolStripStatusLabel2.Text = "                                                               "
        '
        'tssem
        '
        Me.tssem.Name = "tssem"
        Me.tssem.Size = New System.Drawing.Size(0, 17)
        Me.tssem.ToolTipText = "Current Semester"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(12, 17)
        Me.ToolStripStatusLabel1.Text = "-"
        '
        'tsyear
        '
        Me.tsyear.Name = "tsyear"
        Me.tsyear.Size = New System.Drawing.Size(0, 17)
        Me.tsyear.ToolTipText = "Current year"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tstat1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 409)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(881, 22)
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tstat1
        '
        Me.tstat1.Name = "tstat1"
        Me.tstat1.Size = New System.Drawing.Size(38, 17)
        Me.tstat1.Text = "status"
        Me.tstat1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.tstat1.ToolTipText = "Connection Status"
        '
        'DSEGMAIN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(881, 453)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Name = "DSEGMAIN"
        Me.Text = "DSEGSURVEY MAIN"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents tsstatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents menuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuManage As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuUser As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuLout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuLin As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuNS As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuManagesem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tstat1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents menuNA As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuCES As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuGS As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuFRS As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tssem As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsyear As System.Windows.Forms.ToolStripStatusLabel

End Class
