<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmfrsreports
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cbosy = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbocol = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dg1 = New System.Windows.Forms.DataGridView()
        Me.cboryear = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.lblgtot = New System.Windows.Forms.Label()
        Me.bgworker = New System.ComponentModel.BackgroundWorker()
        Me.pbar1 = New System.Windows.Forms.ProgressBar()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblbarep = New System.Windows.Forms.Label()
        Me.btncancelthread = New System.Windows.Forms.Button()
        Me.btnexport = New System.Windows.Forms.Button()
        CType(Me.dg1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cbosy
        '
        Me.cbosy.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cbosy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbosy.FormattingEnabled = True
        Me.cbosy.Location = New System.Drawing.Point(126, 9)
        Me.cbosy.Name = "cbosy"
        Me.cbosy.Size = New System.Drawing.Size(195, 21)
        Me.cbosy.TabIndex = 23
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(22, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 13)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "Select School Year"
        '
        'cbocol
        '
        Me.cbocol.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cbocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbocol.FormattingEnabled = True
        Me.cbocol.Location = New System.Drawing.Point(650, 12)
        Me.cbocol.Name = "cbocol"
        Me.cbocol.Size = New System.Drawing.Size(282, 21)
        Me.cbocol.TabIndex = 25
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(569, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Select College"
        '
        'dg1
        '
        Me.dg1.AllowUserToAddRows = False
        Me.dg1.AllowUserToDeleteRows = False
        Me.dg1.AllowUserToResizeColumns = False
        Me.dg1.AllowUserToResizeRows = False
        Me.dg1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dg1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg1.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dg1.DefaultCellStyle = DataGridViewCellStyle1
        Me.dg1.Location = New System.Drawing.Point(12, 39)
        Me.dg1.Name = "dg1"
        Me.dg1.RowHeadersVisible = False
        Me.dg1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dg1.RowTemplate.Height = 42
        Me.dg1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dg1.Size = New System.Drawing.Size(947, 325)
        Me.dg1.TabIndex = 26
        '
        'cboryear
        '
        Me.cboryear.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cboryear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboryear.FormattingEnabled = True
        Me.cboryear.Items.AddRange(New Object() {"1", "2"})
        Me.cboryear.Location = New System.Drawing.Point(409, 12)
        Me.cboryear.Name = "cboryear"
        Me.cboryear.Size = New System.Drawing.Size(141, 21)
        Me.cboryear.TabIndex = 28
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(341, 15)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Select Year"
        '
        'lblgtot
        '
        Me.lblgtot.AutoSize = True
        Me.lblgtot.Location = New System.Drawing.Point(171, 370)
        Me.lblgtot.Name = "lblgtot"
        Me.lblgtot.Size = New System.Drawing.Size(0, 13)
        Me.lblgtot.TabIndex = 39
        '
        'bgworker
        '
        Me.bgworker.WorkerReportsProgress = True
        Me.bgworker.WorkerSupportsCancellation = True
        '
        'pbar1
        '
        Me.pbar1.Location = New System.Drawing.Point(12, 386)
        Me.pbar1.Name = "pbar1"
        Me.pbar1.Size = New System.Drawing.Size(947, 21)
        Me.pbar1.TabIndex = 40
        Me.pbar1.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 370)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(153, 13)
        Me.Label4.TabIndex = 38
        Me.Label4.Text = "Total Number of Respondent : "
        '
        'lblbarep
        '
        Me.lblbarep.AutoSize = True
        Me.lblbarep.Location = New System.Drawing.Point(506, 370)
        Me.lblbarep.Name = "lblbarep"
        Me.lblbarep.Size = New System.Drawing.Size(44, 13)
        Me.lblbarep.TabIndex = 41
        Me.lblbarep.Text = "lblbarep"
        Me.lblbarep.Visible = False
        '
        'btncancelthread
        '
        Me.btncancelthread.Location = New System.Drawing.Point(815, 363)
        Me.btncancelthread.Name = "btncancelthread"
        Me.btncancelthread.Size = New System.Drawing.Size(144, 26)
        Me.btncancelthread.TabIndex = 42
        Me.btncancelthread.Text = "Cancel Thread 1"
        Me.btncancelthread.UseVisualStyleBackColor = True
        Me.btncancelthread.Visible = False
        '
        'btnexport
        '
        Me.btnexport.Location = New System.Drawing.Point(12, 410)
        Me.btnexport.Name = "btnexport"
        Me.btnexport.Size = New System.Drawing.Size(947, 26)
        Me.btnexport.TabIndex = 43
        Me.btnexport.Text = "Export to Excel"
        Me.btnexport.UseVisualStyleBackColor = True
        '
        'frmfrsreports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(969, 439)
        Me.Controls.Add(Me.btnexport)
        Me.Controls.Add(Me.btncancelthread)
        Me.Controls.Add(Me.lblbarep)
        Me.Controls.Add(Me.pbar1)
        Me.Controls.Add(Me.lblgtot)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboryear)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dg1)
        Me.Controls.Add(Me.cbocol)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbosy)
        Me.Controls.Add(Me.Label2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmfrsreports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Freshmen Retention Survey"
        CType(Me.dg1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbosy As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbocol As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dg1 As System.Windows.Forms.DataGridView
    Friend WithEvents cboryear As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblgtot As System.Windows.Forms.Label
    Friend WithEvents bgworker As System.ComponentModel.BackgroundWorker
    Friend WithEvents pbar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblbarep As System.Windows.Forms.Label
    Friend WithEvents btncancelthread As System.Windows.Forms.Button
    Friend WithEvents btnexport As System.Windows.Forms.Button
End Class
