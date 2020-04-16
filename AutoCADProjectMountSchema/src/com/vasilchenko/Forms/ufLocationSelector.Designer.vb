<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ufLocationSelector
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
        Me.lvPlaces = New System.Windows.Forms.ListView()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lvPlaces
        '
        Me.lvPlaces.HideSelection = False
        Me.lvPlaces.Location = New System.Drawing.Point(12, 12)
        Me.lvPlaces.Name = "lvPlaces"
        Me.lvPlaces.Size = New System.Drawing.Size(500, 200)
        Me.lvPlaces.TabIndex = 5
        Me.lvPlaces.UseCompatibleStateImageBehavior = False
        '
        'btnApply
        '
        Me.btnApply.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnApply.Location = New System.Drawing.Point(356, 226)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(75, 23)
        Me.btnApply.TabIndex = 3
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(437, 226)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'ufLocationSelector
        '
        Me.AcceptButton = Me.btnApply
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(525, 264)
        Me.Controls.Add(Me.lvPlaces)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.btnCancel)
        Me.Name = "ufLocationSelector"
        Me.Text = "LocationSelector"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lvPlaces As Forms.ListView
    Friend WithEvents btnApply As Forms.Button
    Friend WithEvents btnCancel As Forms.Button
End Class
