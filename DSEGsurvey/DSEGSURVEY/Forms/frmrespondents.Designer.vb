<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmrespondents
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgres = New System.Windows.Forms.DataGridView()
        Me.CATNO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RespondentName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Course = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Year = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Gender = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Age = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Religion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Abused = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.AbusedRes = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SEMYEAR = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Selects = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.lblsno = New System.Windows.Forms.Label()
        Me.lbls = New System.Windows.Forms.Label()
        Me.txtsearch = New System.Windows.Forms.TextBox()
        CType(Me.dgres, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgres
        '
        Me.dgres.AllowUserToAddRows = False
        Me.dgres.AllowUserToDeleteRows = False
        Me.dgres.AllowUserToResizeColumns = False
        Me.dgres.AllowUserToResizeRows = False
        Me.dgres.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgres.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dgres.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgres.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CATNO, Me.RespondentName, Me.Course, Me.Year, Me.Gender, Me.Age, Me.Religion, Me.Abused, Me.AbusedRes, Me.SEMYEAR, Me.Selects})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgres.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgres.Location = New System.Drawing.Point(4, 46)
        Me.dgres.Name = "dgres"
        Me.dgres.ReadOnly = True
        Me.dgres.RowHeadersVisible = False
        Me.dgres.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgres.RowTemplate.Height = 42
        Me.dgres.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgres.Size = New System.Drawing.Size(1027, 378)
        Me.dgres.TabIndex = 6
        '
        'CATNO
        '
        Me.CATNO.HeaderText = "Respondent No"
        Me.CATNO.Name = "CATNO"
        Me.CATNO.ReadOnly = True
        Me.CATNO.Width = 70
        '
        'RespondentName
        '
        Me.RespondentName.HeaderText = "RespondentName"
        Me.RespondentName.Name = "RespondentName"
        Me.RespondentName.ReadOnly = True
        Me.RespondentName.Width = 220
        '
        'Course
        '
        Me.Course.HeaderText = "Course"
        Me.Course.Name = "Course"
        Me.Course.ReadOnly = True
        Me.Course.Width = 70
        '
        'Year
        '
        Me.Year.HeaderText = "Year"
        Me.Year.Name = "Year"
        Me.Year.ReadOnly = True
        Me.Year.Width = 50
        '
        'Gender
        '
        Me.Gender.HeaderText = "Gender"
        Me.Gender.Name = "Gender"
        Me.Gender.ReadOnly = True
        Me.Gender.Width = 70
        '
        'Age
        '
        Me.Age.HeaderText = "Age"
        Me.Age.Name = "Age"
        Me.Age.ReadOnly = True
        Me.Age.Width = 30
        '
        'Religion
        '
        Me.Religion.HeaderText = "Religion"
        Me.Religion.Name = "Religion"
        Me.Religion.ReadOnly = True
        Me.Religion.Width = 110
        '
        'Abused
        '
        Me.Abused.HeaderText = "Abused Yes/No"
        Me.Abused.Name = "Abused"
        Me.Abused.ReadOnly = True
        Me.Abused.Width = 50
        '
        'AbusedRes
        '
        Me.AbusedRes.HeaderText = "Abused"
        Me.AbusedRes.Name = "AbusedRes"
        Me.AbusedRes.ReadOnly = True
        Me.AbusedRes.Width = 120
        '
        'SEMYEAR
        '
        Me.SEMYEAR.HeaderText = "Semester"
        Me.SEMYEAR.Name = "SEMYEAR"
        Me.SEMYEAR.ReadOnly = True
        Me.SEMYEAR.Width = 150
        '
        'Selects
        '
        Me.Selects.HeaderText = "SELECT"
        Me.Selects.Name = "Selects"
        Me.Selects.ReadOnly = True
        Me.Selects.Width = 60
        '
        'lblsno
        '
        Me.lblsno.AutoSize = True
        Me.lblsno.Location = New System.Drawing.Point(12, 46)
        Me.lblsno.Name = "lblsno"
        Me.lblsno.Size = New System.Drawing.Size(34, 13)
        Me.lblsno.TabIndex = 30
        Me.lblsno.Text = "lblsno"
        '
        'lbls
        '
        Me.lbls.AutoSize = True
        Me.lbls.Location = New System.Drawing.Point(12, 19)
        Me.lbls.Name = "lbls"
        Me.lbls.Size = New System.Drawing.Size(143, 13)
        Me.lbls.TabIndex = 31
        Me.lbls.Text = "Search Respondent Details :"
        '
        'txtsearch
        '
        Me.txtsearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtsearch.Location = New System.Drawing.Point(161, 17)
        Me.txtsearch.Name = "txtsearch"
        Me.txtsearch.Size = New System.Drawing.Size(581, 20)
        Me.txtsearch.TabIndex = 32
        '
        'frmrespondents
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1034, 436)
        Me.Controls.Add(Me.txtsearch)
        Me.Controls.Add(Me.lbls)
        Me.Controls.Add(Me.dgres)
        Me.Controls.Add(Me.lblsno)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmrespondents"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "RESPONDENT"
        CType(Me.dgres, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dgres As System.Windows.Forms.DataGridView
    Friend WithEvents lblsno As System.Windows.Forms.Label
    Friend WithEvents CATNO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents RespondentName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Course As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Year As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Gender As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Age As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Religion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Abused As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AbusedRes As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SEMYEAR As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Selects As System.Windows.Forms.DataGridViewButtonColumn
    Friend WithEvents lbls As System.Windows.Forms.Label
    Friend WithEvents txtsearch As System.Windows.Forms.TextBox
End Class
