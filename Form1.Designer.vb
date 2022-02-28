<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.btn_OpenPoint = New System.Windows.Forms.Button()
        Me.logFrame = New System.Windows.Forms.TextBox()
        Me.txt_SSTache = New System.Windows.Forms.TextBox()
        Me.txt_Tache = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'btn_OpenPoint
        '
        Me.btn_OpenPoint.Location = New System.Drawing.Point(11, 12)
        Me.btn_OpenPoint.Name = "btn_OpenPoint"
        Me.btn_OpenPoint.Size = New System.Drawing.Size(230, 47)
        Me.btn_OpenPoint.TabIndex = 0
        Me.btn_OpenPoint.Text = "Pointage"
        Me.btn_OpenPoint.UseVisualStyleBackColor = True
        '
        'logFrame
        '
        Me.logFrame.Location = New System.Drawing.Point(11, 153)
        Me.logFrame.Multiline = True
        Me.logFrame.Name = "logFrame"
        Me.logFrame.Size = New System.Drawing.Size(230, 88)
        Me.logFrame.TabIndex = 2
        '
        'txt_SSTache
        '
        Me.txt_SSTache.Location = New System.Drawing.Point(12, 124)
        Me.txt_SSTache.Name = "txt_SSTache"
        Me.txt_SSTache.Size = New System.Drawing.Size(229, 23)
        Me.txt_SSTache.TabIndex = 3
        '
        'txt_Tache
        '
        Me.txt_Tache.Location = New System.Drawing.Point(12, 80)
        Me.txt_Tache.Name = "txt_Tache"
        Me.txt_Tache.Size = New System.Drawing.Size(229, 23)
        Me.txt_Tache.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 15)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Tâche :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 106)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 15)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Sous-tâche :"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(249, 247)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_Tache)
        Me.Controls.Add(Me.txt_SSTache)
        Me.Controls.Add(Me.logFrame)
        Me.Controls.Add(Me.btn_OpenPoint)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AutoPoint"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_OpenPoint As Button
    Friend WithEvents logFrame As TextBox
    Friend WithEvents txt_SSTache As TextBox
    Friend WithEvents txt_Tache As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Public WithEvents Timer1 As Timer
End Class
