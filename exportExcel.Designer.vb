<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class exportExcel
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.chbx_lighting = New System.Windows.Forms.CheckBox()
        Me.chbx_screen = New System.Windows.Forms.CheckBox()
        Me.chbx_truss = New System.Windows.Forms.CheckBox()
        Me.chbx_constr = New System.Windows.Forms.CheckBox()
        Me.chbx_sound = New System.Windows.Forms.CheckBox()
        Me.chbx_commutation = New System.Windows.Forms.CheckBox()
        Me.chbx_all = New System.Windows.Forms.CheckBox()
        Me.btn_export = New System.Windows.Forms.Button()
        Me.btn_cancelExport = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'chbx_lighting
        '
        Me.chbx_lighting.AutoSize = True
        Me.chbx_lighting.Location = New System.Drawing.Point(71, 58)
        Me.chbx_lighting.Name = "chbx_lighting"
        Me.chbx_lighting.Size = New System.Drawing.Size(63, 17)
        Me.chbx_lighting.TabIndex = 0
        Me.chbx_lighting.Text = "Lighting"
        Me.chbx_lighting.UseVisualStyleBackColor = True
        '
        'chbx_screen
        '
        Me.chbx_screen.AutoSize = True
        Me.chbx_screen.Location = New System.Drawing.Point(71, 81)
        Me.chbx_screen.Name = "chbx_screen"
        Me.chbx_screen.Size = New System.Drawing.Size(60, 17)
        Me.chbx_screen.TabIndex = 0
        Me.chbx_screen.Text = "Screen"
        Me.chbx_screen.UseVisualStyleBackColor = True
        '
        'chbx_truss
        '
        Me.chbx_truss.AutoSize = True
        Me.chbx_truss.Location = New System.Drawing.Point(71, 127)
        Me.chbx_truss.Name = "chbx_truss"
        Me.chbx_truss.Size = New System.Drawing.Size(107, 17)
        Me.chbx_truss.TabIndex = 0
        Me.chbx_truss.Text = "Truss and motors"
        Me.chbx_truss.UseVisualStyleBackColor = True
        '
        'chbx_constr
        '
        Me.chbx_constr.AutoSize = True
        Me.chbx_constr.Location = New System.Drawing.Point(71, 150)
        Me.chbx_constr.Name = "chbx_constr"
        Me.chbx_constr.Size = New System.Drawing.Size(85, 17)
        Me.chbx_constr.TabIndex = 0
        Me.chbx_constr.Text = "Construction"
        Me.chbx_constr.UseVisualStyleBackColor = True
        '
        'chbx_sound
        '
        Me.chbx_sound.AutoSize = True
        Me.chbx_sound.Location = New System.Drawing.Point(71, 173)
        Me.chbx_sound.Name = "chbx_sound"
        Me.chbx_sound.Size = New System.Drawing.Size(57, 17)
        Me.chbx_sound.TabIndex = 0
        Me.chbx_sound.Text = "Sound"
        Me.chbx_sound.UseVisualStyleBackColor = True
        '
        'chbx_commutation
        '
        Me.chbx_commutation.AutoSize = True
        Me.chbx_commutation.Location = New System.Drawing.Point(71, 104)
        Me.chbx_commutation.Name = "chbx_commutation"
        Me.chbx_commutation.Size = New System.Drawing.Size(87, 17)
        Me.chbx_commutation.TabIndex = 0
        Me.chbx_commutation.Text = "Commutation"
        Me.chbx_commutation.UseVisualStyleBackColor = True
        '
        'chbx_all
        '
        Me.chbx_all.AutoSize = True
        Me.chbx_all.Location = New System.Drawing.Point(48, 29)
        Me.chbx_all.Name = "chbx_all"
        Me.chbx_all.Size = New System.Drawing.Size(37, 17)
        Me.chbx_all.TabIndex = 1
        Me.chbx_all.Text = "All"
        Me.chbx_all.UseVisualStyleBackColor = True
        '
        'btn_export
        '
        Me.btn_export.Location = New System.Drawing.Point(31, 221)
        Me.btn_export.Name = "btn_export"
        Me.btn_export.Size = New System.Drawing.Size(75, 23)
        Me.btn_export.TabIndex = 2
        Me.btn_export.Text = "Export"
        Me.btn_export.UseVisualStyleBackColor = True
        '
        'btn_cancelExport
        '
        Me.btn_cancelExport.Location = New System.Drawing.Point(129, 220)
        Me.btn_cancelExport.Name = "btn_cancelExport"
        Me.btn_cancelExport.Size = New System.Drawing.Size(75, 23)
        Me.btn_cancelExport.TabIndex = 3
        Me.btn_cancelExport.Text = "Cancel"
        Me.btn_cancelExport.UseVisualStyleBackColor = True
        '
        'exportExcel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(250, 273)
        Me.Controls.Add(Me.btn_cancelExport)
        Me.Controls.Add(Me.btn_export)
        Me.Controls.Add(Me.chbx_all)
        Me.Controls.Add(Me.chbx_commutation)
        Me.Controls.Add(Me.chbx_sound)
        Me.Controls.Add(Me.chbx_constr)
        Me.Controls.Add(Me.chbx_truss)
        Me.Controls.Add(Me.chbx_screen)
        Me.Controls.Add(Me.chbx_lighting)
        Me.Name = "exportExcel"
        Me.Text = "exportExcel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents chbx_lighting As CheckBox
    Friend WithEvents chbx_screen As CheckBox
    Friend WithEvents chbx_truss As CheckBox
    Friend WithEvents chbx_constr As CheckBox
    Friend WithEvents chbx_sound As CheckBox
    Friend WithEvents chbx_commutation As CheckBox
    Friend WithEvents chbx_all As CheckBox
    Friend WithEvents btn_export As Button
    Friend WithEvents btn_cancelExport As Button
End Class
