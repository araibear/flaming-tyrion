<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button5 = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button8 = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(700, 545)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(25, 255)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(86, 24)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "消去"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(25, 285)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(86, 24)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "終了"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(774, 739)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(75, 19)
        Me.TextBox2.TabIndex = 5
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(864, 739)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(75, 19)
        Me.TextBox3.TabIndex = 6
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(368, 11)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(134, 30)
        Me.Button4.TabIndex = 7
        Me.Button4.Text = "検索"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("ＭＳ 明朝", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(104, 13)
        Me.TextBox1.MaxLength = 1
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(29, 23)
        Me.TextBox1.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("ＭＳ 明朝", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(22, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 16)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "検索文字"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("ＭＳ 明朝", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(192, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 16)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "被験者番号"
        '
        'TextBox4
        '
        Me.TextBox4.Font = New System.Drawing.Font("ＭＳ 明朝", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(291, 13)
        Me.TextBox4.MaxLength = 3
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(46, 23)
        Me.TextBox4.TabIndex = 12
        '
        'PictureBox2
        '
        Me.PictureBox2.Location = New System.Drawing.Point(3, -7)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(1400, 77)
        Me.PictureBox2.TabIndex = 13
        Me.PictureBox2.TabStop = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(10, 22)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(54, 16)
        Me.CheckBox1.TabIndex = 14
        Me.CheckBox1.Text = "1回目"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Location = New System.Drawing.Point(10, 42)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(54, 16)
        Me.CheckBox2.TabIndex = 15
        Me.CheckBox2.Text = "2回目"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Location = New System.Drawing.Point(11, 63)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(54, 16)
        Me.CheckBox3.TabIndex = 16
        Me.CheckBox3.Text = "3回目"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Location = New System.Drawing.Point(12, 85)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(54, 16)
        Me.CheckBox4.TabIndex = 17
        Me.CheckBox4.Text = "4回目"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Location = New System.Drawing.Point(13, 107)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(54, 16)
        Me.CheckBox5.TabIndex = 18
        Me.CheckBox5.Text = "5回目"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 131)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(54, 19)
        Me.Button1.TabIndex = 19
        Me.Button1.Text = "全選択"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(65, 132)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(49, 18)
        Me.Button6.TabIndex = 20
        Me.Button6.Text = "クリア"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.CheckBox5)
        Me.GroupBox1.Controls.Add(Me.CheckBox4)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 51)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(116, 161)
        Me.GroupBox1.TabIndex = 21
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "回数選択"
        '
        'CheckBox6
        '
        Me.CheckBox6.AutoSize = True
        Me.CheckBox6.Location = New System.Drawing.Point(23, 228)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(66, 16)
        Me.CheckBox6.TabIndex = 22
        Me.CheckBox6.Text = "筆順あり"
        Me.CheckBox6.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(132, 850)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(47, 12)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "アクション"
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 12
        Me.ListBox1.Location = New System.Drawing.Point(12, 696)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(97, 28)
        Me.ListBox1.TabIndex = 24
        Me.ListBox1.Visible = False
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Location = New System.Drawing.Point(132, 51)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(704, 551)
        Me.Panel1.TabIndex = 25
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(544, 11)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(118, 30)
        Me.Button5.TabIndex = 26
        Me.Button5.Text = "印刷"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.AutoScroll = True
        Me.Panel2.Controls.Add(Me.PictureBox2)
        Me.Panel2.Location = New System.Drawing.Point(111, 604)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(780, 87)
        Me.Panel2.TabIndex = 27
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(723, 11)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(29, 24)
        Me.Button7.TabIndex = 28
        Me.Button7.Text = "<"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(804, 13)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(30, 23)
        Me.Button8.TabIndex = 29
        Me.Button8.Text = ">"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Menu
        Me.Label4.Location = New System.Drawing.Point(772, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(17, 12)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "01"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(952, 714)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CheckBox6)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Name = "Form1"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Text = "筆跡検索システム"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label

End Class
