<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmchangesem
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmchangesem))
        Me.dgsem = New System.Windows.Forms.DataGridView()
        Me.SID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Category = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.mYEAR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.answers = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.sremove = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbosem = New System.Windows.Forms.ComboBox()
        Me.cboyear = New System.Windows.Forms.ComboBox()
        Me.tsot1 = New System.Windows.Forms.ToolStrip()
        Me.tsrefresh = New System.Windows.Forms.ToolStripButton()
        Me.saveme = New System.Windows.Forms.ToolStripButton()
        CType(Me.dgsem, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tsot1.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgsem
        '
        Me.dgsem.AllowUserToAddRows = False
        Me.dgsem.AllowUserToDeleteRows = False
        Me.dgsem.AllowUserToResizeColumns = False
        Me.dgsem.AllowUserToResizeRows = False
        Me.dgsem.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgsem.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dgsem.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgsem.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.SID, Me.Category, Me.mYEAR, Me.answers, Me.sremove})
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgsem.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgsem.Location = New System.Drawing.Point(9, 79)
        Me.dgsem.Name = "dgsem"
        Me.dgsem.RowHeadersVisible = False
        Me.dgsem.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgsem.RowTemplate.Height = 42
        Me.dgsem.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgsem.Size = New System.Drawing.Size(503, 263)
        Me.dgsem.TabIndex = 20
        '
        'SID
        '
        Me.SID.HeaderText = "SID"
        Me.SID.Name = "SID"
        Me.SID.Visible = False
        '
        'Category
        '
        Me.Category.HeaderText = "SEMESTER"
        Me.Category.Name = "Category"
        Me.Category.Width = 200
        '
        'mYEAR
        '
        Me.mYEAR.HeaderText = "YEAR"
        Me.mYEAR.Name = "mYEAR"
        '
        'answers
        '
        Me.answers.DataPropertyName = "answers"
        Me.answers.HeaderText = "CHOOSE"
        Me.answers.Name = "answers"
        Me.answers.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.answers.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.answers.ToolTipText = "click cell to choose"
        '
        'sremove
        '
        Me.sremove.HeaderText = "REMOVE"
        Me.sremove.Name = "sremove"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(21, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 13)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "SEMESTER"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(259, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "YEAR"
        '
        'cbosem
        '
        Me.cbosem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbosem.FormattingEnabled = True
        Me.cbosem.Items.AddRange(New Object() {"Summer", "First Semester", "Second Semester"})
        Me.cbosem.Location = New System.Drawing.Point(93, 6)
        Me.cbosem.Name = "cbosem"
        Me.cbosem.Size = New System.Drawing.Size(156, 21)
        Me.cbosem.TabIndex = 33
        '
        'cboyear
        '
        Me.cboyear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboyear.FormattingEnabled = True
        Me.cboyear.Location = New System.Drawing.Point(301, 6)
        Me.cboyear.Name = "cboyear"
        Me.cboyear.Size = New System.Drawing.Size(211, 21)
        Me.cboyear.TabIndex = 34
        '
        'tsot1
        '
        Me.tsot1.AutoSize = False
        Me.tsot1.BackColor = System.Drawing.Color.Khaki
        Me.tsot1.Dock = System.Windows.Forms.DockStyle.None
        Me.tsot1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsrefresh, Me.saveme})
        Me.tsot1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow
        Me.tsot1.Location = New System.Drawing.Point(11, 30)
        Me.tsot1.Name = "tsot1"
        Me.tsot1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tsot1.Size = New System.Drawing.Size(503, 46)
        Me.tsot1.TabIndex = 35
        Me.tsot1.Text = "ToolStrip1"
        '
        'tsrefresh
        '
        Me.tsrefresh.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsrefresh.AutoSize = False
        Me.tsrefresh.Image = Global.DSEGSURVEY.My.Resources.Resources._001_06
        Me.tsrefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsrefresh.Name = "tsrefresh"
        Me.tsrefresh.Padding = New System.Windows.Forms.Padding(9)
        Me.tsrefresh.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tsrefresh.Size = New System.Drawing.Size(140, 45)
        Me.tsrefresh.Text = "Refresh"
        Me.tsrefresh.ToolTipText = "Reset Ctrl + R"
        '
        'saveme
        '
        Me.saveme.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.saveme.AutoSize = False
        Me.saveme.Image = CType(resources.GetObject("saveme.Image"), System.Drawing.Image)
        Me.saveme.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.saveme.Name = "saveme"
        Me.saveme.Padding = New System.Windows.Forms.Padding(9)
        Me.saveme.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveme.Size = New System.Drawing.Size(140, 45)
        Me.saveme.Text = "Add New Semester "
        '
        'frmchangesem
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(523, 350)
        Me.Controls.Add(Me.tsot1)
        Me.Controls.Add(Me.cboyear)
        Me.Controls.Add(Me.cbosem)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgsem)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmchangesem"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CHANGE SEMESTER"
        CType(Me.dgsem,System.ComponentModel.ISupportInitialize).EndInit
        Me.tsot1.ResumeLayout(false)
        Me.tsot1.PerformLayout
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents dgsem As System.Windows.Forms.DataGridView
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbosem As System.Windows.Forms.ComboBox
    Friend WithEvents cboyear As System.Windows.Forms.ComboBox
    Friend WithEvents tsot1 As System.Windows.Forms.ToolStrip
    Friend WithEvents tsrefresh As System.Windows.Forms.ToolStripButton
    Friend WithEvents saveme As System.Windows.Forms.ToolStripButton
    Friend WithEvents SID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Category As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents mYEAR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents answers As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents sremove As System.Windows.Forms.DataGridViewButtonColumn
End Class
