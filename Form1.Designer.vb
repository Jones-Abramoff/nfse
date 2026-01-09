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
        Me.Msg = New System.Windows.Forms.ListBox()
        Me.Lote = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Msg
        '
        Me.Msg.FormattingEnabled = True
        Me.Msg.HorizontalExtent = 2000
        Me.Msg.HorizontalScrollbar = True
        Me.Msg.Location = New System.Drawing.Point(12, 107)
        Me.Msg.Name = "Msg"
        Me.Msg.Size = New System.Drawing.Size(928, 225)
        Me.Msg.TabIndex = 11
        '
        'Lote
        '
        Me.Lote.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lote.Location = New System.Drawing.Point(59, 20)
        Me.Lote.Name = "Lote"
        Me.Lote.Size = New System.Drawing.Size(58, 18)
        Me.Lote.TabIndex = 10
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 60)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(928, 26)
        Me.ProgressBar1.TabIndex = 8
        '
        'Timer1
        '
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(22, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Lote:"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(952, 356)
        Me.Controls.Add(Me.Msg)
        Me.Controls.Add(Me.Lote)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "CORPORATOR - Log de Nota Fiscal de Serviço Eletrônica - 20260109"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Msg As System.Windows.Forms.ListBox
    Friend WithEvents Lote As System.Windows.Forms.Label
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
