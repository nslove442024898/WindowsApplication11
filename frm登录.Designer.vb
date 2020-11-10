<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm登录
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblmsg = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.tbxPWD = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxUsername = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.bntlogin = New System.Windows.Forms.Button()
        Me.Timer余额查询时钟 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(305, 171)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(0, 12)
        Me.Label4.TabIndex = 18
        '
        'lblmsg
        '
        Me.lblmsg.AutoSize = True
        Me.lblmsg.Location = New System.Drawing.Point(25, 111)
        Me.lblmsg.Name = "lblmsg"
        Me.lblmsg.Size = New System.Drawing.Size(0, 12)
        Me.lblmsg.TabIndex = 17
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RichTextBox1.BulletIndent = 10
        Me.RichTextBox1.Location = New System.Drawing.Point(291, 12)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(389, 156)
        Me.RichTextBox1.TabIndex = 15
        Me.RichTextBox1.Text = "软件使用注意事项:" & Global.Microsoft.VisualBasic.ChrW(10) & "1.请输入您的用户名和密码以使用本软件；" & Global.Microsoft.VisualBasic.ChrW(10) & "2.若您尚未取得用户名，请退出本软件，使用您的电脑访问网站：https:\\www.TopYantu" &
    ".tech，在网站注册您的用户名。然后将用户名告诉QQ40469586，我们将给您的账号充值，以便近期内侧人员的调试。" & Global.Microsoft.VisualBasic.ChrW(10) & "3.若使用过程中出现软件故障，可向QQ4" &
    "0469586说明，以便修正。" & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'tbxPWD
        '
        Me.tbxPWD.Location = New System.Drawing.Point(83, 61)
        Me.tbxPWD.Name = "tbxPWD"
        Me.tbxPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbxPWD.Size = New System.Drawing.Size(151, 21)
        Me.tbxPWD.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(30, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 12)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "密  码"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(30, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(41, 12)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "用户名"
        '
        'tbxUsername
        '
        Me.tbxUsername.Location = New System.Drawing.Point(83, 27)
        Me.tbxUsername.Name = "tbxUsername"
        Me.tbxUsername.Size = New System.Drawing.Size(151, 21)
        Me.tbxUsername.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.tbxPWD)
        Me.GroupBox1.Controls.Add(Me.tbxUsername)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(273, 103)
        Me.GroupBox1.TabIndex = 20
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "请输入用户名和密码"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(10, 239)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(0, 12)
        Me.lblStatus.TabIndex = 21
        '
        'bntlogin
        '
        Me.bntlogin.Location = New System.Drawing.Point(91, 132)
        Me.bntlogin.Name = "bntlogin"
        Me.bntlogin.Size = New System.Drawing.Size(155, 62)
        Me.bntlogin.TabIndex = 22
        Me.bntlogin.Text = "登录"
        Me.bntlogin.UseVisualStyleBackColor = True
        '
        'frm登录
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(710, 294)
        Me.Controls.Add(Me.bntlogin)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblmsg)
        Me.Controls.Add(Me.RichTextBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm登录"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "登录"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label4 As Label
    Friend WithEvents lblmsg As Label
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents tbxPWD As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxUsername As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents bntlogin As Button
    Friend WithEvents Timer余额查询时钟 As Timer
End Class
