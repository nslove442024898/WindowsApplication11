Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.IO

'Imports 前沿数据库工具独立版V

Public Class frm登录

    Public WithEvents c As Cls用户信息和费用余额查询扣减器
    Dim WithEvents frmWork As Form1
    Dim LuserInfo As List(Of String)

    'Dim feeRate As Double = 2



    Dim Dir当前用户的记账文件夹名称 As String

    Dim lbl用户占用标志文件 As String
    Dim CheckLoginTime As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles bntlogin.Click
        Try
            frmWork = New Form1()
            frmWork.Show()
            Exit Sub
            If tbxUsername.Text = "" Or tbxPWD.Text = "" Then
                MsgBox("用户名和密码不能为空")
                Exit Sub
            End If

            If lbl用户占用标志文件 = "" Then
                lbl用户占用标志文件 = "C:\er\" & tbxUsername.Text & "\" & tbxUsername.Text & "正在使用数据库工具软件.txt"
                Dir当前用户的记账文件夹名称 = tbxUsername.Text
                If My.Computer.FileSystem.DirectoryExists("c\er\" & Dir当前用户的记账文件夹名称) = False Then
                    My.Computer.FileSystem.CreateDirectory("c:\er\" & Dir当前用户的记账文件夹名称)
                End If
                If File.Exists(lbl用户占用标志文件) Then
                    MsgBox("您输入的用户名正在使用中，请输入其他用户名。")
                    Me.Show()
                    Exit Sub
                End If
                Dim s As FileStream = File.Create(lbl用户占用标志文件)
                s.Close()
            Else
                If lbl用户占用标志文件 = "C:\er\" & tbxUsername.Text & "\" & tbxUsername.Text & "正在使用数据库工具软件.txt" And File.Exists(lbl用户占用标志文件) Then
                    MsgBox("您输入的用户名正在使用中，请输入其他用户名。")
                    Me.Show()
                    Exit Sub
                Else
                    If My.Computer.FileSystem.DirectoryExists("c\er\" & Dir当前用户的记账文件夹名称) = False Then
                        My.Computer.FileSystem.CreateDirectory("c:\er\" & Dir当前用户的记账文件夹名称)
                    End If
                    Dim s As FileStream = File.Create(lbl用户占用标志文件)
                    s.Close()
                End If
            End If
            '此处根据软件修改====================================================
            c = New Cls用户信息和费用余额查询扣减器(tbxUsername.Text, tbxPWD.Text, 0.5, "databasetool")

            c.连接_构造用户_查询该用户余额并引发欠费事件()

        Catch ex As Exception
            MsgBox(Err.Description)
            Try
                File.Delete(lbl用户占用标志文件)
            Catch ex1 As Exception

            End Try
        End Try

    End Sub






    Private Sub tbxUsername_TextChanged(sender As Object, e As EventArgs) Handles tbxUsername.TextChanged

    End Sub

    Private Sub tbxUsername_KeyDown(sender As Object, e As KeyEventArgs) Handles tbxUsername.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If tbxPWD.Text = "" Or tbxUsername.Text = "" Then
                    MsgBox("用户名和密码不能为空")
                    Exit Sub
                Else
                    bntLogin.PerformClick()
                End If

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tbxPWD_KeyDown(sender As Object, e As KeyEventArgs) Handles tbxPWD.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then
                If tbxPWD.Text = "" Or tbxUsername.Text = "" Then
                    MsgBox("用户名和密码不能为空")
                    Exit Sub
                Else
                    bntLogin.PerformClick()
                End If

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub frm登录_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxUsername.Select()
    End Sub





    Private Sub c_用户状态改变了(ByRef e As enumstatus用户状态) Handles c.用户状态改变了
        Try
            Select Case e
                Case enumstatus用户状态.正常
                    frmWork = New Form1
                    frmWork.PubCurrentUserName当前用户 = c.CurUser当前用户.username
                    frmWork.Show()
                    frmWork.lblremainningtime.Text = "剩余时间：" & c.当前用户剩余时间
                    frmWork.lblremainningtime.Refresh()
                    Dim str As String = "d:\rcs\用户使用记录.txt"
                    File.AppendAllText(str, c.CurUser当前用户.username & "," & Now.ToString & "," & System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Namespace & vbCrLf)
                    Me.Visible = False
                Case enumstatus用户状态.费用低
                    If IsNothing(Me.frmWork) Then
                        frmWork.lblStatus.Text = "用户费用低，请充值。"
                        frmWork.lblStatus.Refresh()

                    End If
                Case enumstatus用户状态.欠费 Or enumstatus用户状态.过期
                    lblStatus.Text = "用户已欠费或过期"
                    frmWork.Close()
                    Me.Visible = True
                Case enumstatus用户状态.用户不存在
                    lblStatus.Text = "用户不存在或密码错误"

            End Select
        Catch ex As Exception
            MsgBox("用户状态改变过程出错" & Err.Description)
        End Try
    End Sub

    Private Sub frmWork_Closed(sender As Object, e As EventArgs) Handles frmWork.Closed
        If File.Exists(lbl用户占用标志文件) Then File.Delete(lbl用户占用标志文件)
        lbl用户占用标志文件 = ""
        Try
            If Dir当前用户的记账文件夹名称 <> "" Then My.Computer.FileSystem.DeleteDirectory("c:\er\" & Dir当前用户的记账文件夹名称, FileIO.DeleteDirectoryOption.DeleteAllContents)
        Catch ex As Exception

        End Try
        Me.Close()
    End Sub

    Private Sub c_查无此用户() Handles c.查无此用户
        Try
            lblStatus.Text = "用户名不存在或密码错误。"
            lblStatus.Refresh()
            If File.Exists(lbl用户占用标志文件) Then File.Delete(lbl用户占用标志文件)
            lbl用户占用标志文件 = ""
        Catch ex As Exception

        End Try

    End Sub

    Private Sub c_有此用户() Handles c.有此用户
        lblStatus.Text = ""
        lblStatus.Refresh()
    End Sub

    Private Sub c_用户余额变化了() Handles c.用户余额变化了
        Try
            frmWork.lblremainningtime.Text = "剩余时间：" & c.当前用户剩余时间
            frmWork.lblremainningtime.Refresh()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub c_用户没有购买本软件() Handles c.用户没有购买本软件
        Try
            lblStatus.Text = "用户没有购买此软件"
            File.Delete(lbl用户占用标志文件)
            lbl用户占用标志文件 = ""
        Catch ex As Exception
            MsgBox("用户没有购买本软件事件出错" & Err.Description)
        End Try
    End Sub

    Private Sub c_此用户为过期或欠费用户() Handles c.此用户为过期或欠费用户
        Try
            lblStatus.Text = "此用户已欠费或过期"
            lblStatus.Refresh()
            If File.Exists(lbl用户占用标志文件) Then File.Delete(lbl用户占用标志文件)
        Catch ex As Exception

        End Try
    End Sub
End Class
