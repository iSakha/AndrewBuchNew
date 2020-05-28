<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class sumForm
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
        Me.dgv_sum = New System.Windows.Forms.DataGridView()
        Me.btn_close = New System.Windows.Forms.Button()
        CType(Me.dgv_sum, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgv_sum
        '
        Me.dgv_sum.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_sum.Location = New System.Drawing.Point(0, 0)
        Me.dgv_sum.Name = "dgv_sum"
        Me.dgv_sum.Size = New System.Drawing.Size(785, 548)
        Me.dgv_sum.TabIndex = 0
        '
        'btn_close
        '
        Me.btn_close.Location = New System.Drawing.Point(355, 566)
        Me.btn_close.Name = "btn_close"
        Me.btn_close.Size = New System.Drawing.Size(75, 23)
        Me.btn_close.TabIndex = 1
        Me.btn_close.Text = "Close"
        Me.btn_close.UseVisualStyleBackColor = True
        '
        'sumForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 601)
        Me.Controls.Add(Me.btn_close)
        Me.Controls.Add(Me.dgv_sum)
        Me.Name = "sumForm"
        Me.Text = "Form2"
        CType(Me.dgv_sum, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgv_sum As DataGridView
    Friend WithEvents btn_close As Button
End Class
