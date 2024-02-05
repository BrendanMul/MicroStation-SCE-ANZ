<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatus
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStatus))
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.lblWait = New System.Windows.Forms.Label()
        Me.prgProgress = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.lblStatus.Location = New System.Drawing.Point(9, 9)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(457, 13)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "Status"
        '
        'lblWait
        '
        Me.lblWait.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.lblWait.AutoSize = True
        Me.lblWait.Location = New System.Drawing.Point(9, 33)
        Me.lblWait.Name = "lblWait"
        Me.lblWait.Size = New System.Drawing.Size(70, 13)
        Me.lblWait.TabIndex = 1
        Me.lblWait.Text = "Please wait..."
        '
        'prgProgress
        '
        Me.prgProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.prgProgress.Location = New System.Drawing.Point(12, 49)
        Me.prgProgress.Name = "prgProgress"
        Me.prgProgress.Size = New System.Drawing.Size(454, 18)
        Me.prgProgress.TabIndex = 2
        Me.prgProgress.Visible = False
        '
        'frmStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(478, 79)
        Me.Controls.Add(Me.prgProgress)
        Me.Controls.Add(Me.lblWait)
        Me.Controls.Add(Me.lblStatus)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmStatus"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Status"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents lblStatus As System.Windows.Forms.Label
    Public WithEvents lblWait As System.Windows.Forms.Label
    Public WithEvents prgProgress As System.Windows.Forms.ProgressBar
End Class
