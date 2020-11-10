Imports System.Collections.Generic
Imports System.IO
Imports com.ms.win32.Kernel32

Public Class Form2
    Public WithEvents tCheckLogin As Timer
    Public WithEvents tCheckUserInfo As Timer
    Public WithEvents c As Checker
    Dim WithEvents frmWork As Form1
    Dim LuserInfo As List(Of String)
    Dim pId As Long
    Dim hProcess As Long
    Dim feeRate As Double
    Dim dir As String
    Dim CheckLoginTime As Integer
    Public Sub New()
        tCheckLogin = New Timer
        tCheckUserInfo = New Timer
        feeRate = 0.5
    End Sub


    Public Sub checkUser()

    End Sub
    Function MakeDir() As Boolean
        Dim i As Integer = 0
        Do While True
            Randomize()
            Dim dirstr As String = Int(1000 * Rnd())
            If isDirExist("c:\er\" & dirstr) = False Then
                My.Computer.FileSystem.CreateDirectory("c:\er\" & dirstr)
                dir = dirstr
                Return True
            Else
                i = i + 1
            End If
            If i > 100 Then
                Return False
            End If
        Loop
    End Function
    Public Function isDirExist(ByVal strPath As String) As Boolean
        Dim strDirTemp As String()
        strDirTemp = strPath.Split("\")
        strPath = String.Empty
        For i As Integer = 0 To strDirTemp.Length - 1
            ' 判断数组内容.目的是防止输入的strPath内容如:c:\abc\123\ 最后一位也是"\"
            If strDirTemp(i) <> "" Then
                strPath += strDirTemp(i) & "\"
            End If
        Next

        ' 判断文件夹是否存在
        isDirExist = System.IO.Directory.Exists(strPath)
    End Function
    Sub checkLogin() Handles tCheckLogin.Tick
        Try
            If File.Exists("c:\er\" & dir & "\QYSJKusername.txt") = False Then
                CheckLoginTime = CheckLoginTime + 1
                If CheckLoginTime > 180 Then
                    MsgBox("用于您长时间未能完成登录，本软件暂时关闭，若您需要再次登录请重新输入命令QYKC")
                    TerminateProcess(hProcess, 3838)

                    Try
                        File.Delete("c:\er\" & dir & "\QYSJKusername.txt")
                        My.Computer.FileSystem.DeleteDirectory("c:\er\" & dir, FileIO.DeleteDirectoryOption.DeleteAllContents)
                        File.Delete("c:\er\" & c.UserInfo.username & ".txt")
                        frmWork = Nothing
                        tCheckLogin.Stop()
                        tCheckLogin.Dispose()
                        tCheckUserInfo.Stop()
                        tCheckUserInfo.Dispose()
                    Catch ex As Exception

                    End Try


                    Me.Finalize()

                End If
                Exit Sub
            End If
            c.userNameDocPath = "c:\er\" & dir & "\QYSJKusername.txt"
            c.ReadUserInfo()

            tCheckLogin.Stop()
            '看账号类型，如果是管理员则启动工作窗体，如果是体验用户或正常用户，如果时间或金钱大于0则其他工作窗体
            Select Case c.UserInfo.userID
                Case UserIdentity.Administrator
                    If IsNothing(Me.frmWork) Then
                        frmWork = New Form1
                        frmWork.Show()
                    Else
                        Me.frmWork.WindowState = FormWindowState.Normal
                    End If


                Case UserIdentity.TrialUser
                    If DateDiff(DateInterval.Minute, Now, CDate(c.UserInfo.validDate)) > 0 Then
                        If IsNothing(Me.frmWork) Then
                            frmWork = New Form1
                            frmWork.Show()
                        Else
                            Me.frmWork.WindowState = FormWindowState.Normal
                        End If
                    Else
                        MsgBox("用户已过期")
                        Exit Sub
                    End If
                Case UserIdentity.User
                    If c.UserInfo.money > 0 Then
                        If IsNothing(Me.frmWork) Then
                            frmWork = New Form1
                            frmWork.Show()
                        Else
                            Me.frmWork.WindowState = FormWindowState.Normal
                        End If
                    Else
                        MsgBox("用户账号余额不足")
                        Exit Sub
                    End If
            End Select
            '查询用户登录进程停止，启动查询用户进程
            tCheckUserInfo.Interval = 6000
            tCheckUserInfo.Start()
            Exit Sub
        Catch ex As Exception
            MsgBox("检查登录用户过程出错" & Err.Description)
        End Try

    End Sub
    Sub checkUserInfo() Handles tCheckUserInfo.Tick
        c.ReadUserInfo()
        If IsNothing(frmWork) = False Then
            frmWork.lblcurUserName.Text = "用户名：" & c.UserInfo.username.ToString
            frmWork.lbluserid.Text = "用户类型：" & c.UserInfo.userID.ToString
            frmWork.lblremainningtime.Text = "剩余时间：" & FormatNumber(c.UserInfo.money / feeRate, 2)
            frmWork.lblExpireTime.Text = "到期时间：" & c.UserInfo.validDate
        End If
    End Sub
    Sub timeOutOrMoneyOUT() Handles c.UserMoneyOut, c.UserTimeOut
        Try
            MsgBox("用户期限已到")
            frmWork.Close()
            tCheckUserInfo.Stop()
        Catch ex As Exception

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        File.Delete("c:\er\" & c.UserInfo.username & ".txt")
        MyBase.Finalize()
    End Sub

    Private Sub frmWork_Closed(sender As Object, e As EventArgs) Handles frmWork.Closed
        TerminateProcess(hProcess, 3838)
        If Not dir <> "" Then
            File.Delete("c:\er\" & dir & "\QYSJKusername.txt")
            My.Computer.FileSystem.DeleteDirectory("c:\er\" & dir, FileIO.DeleteDirectoryOption.DeleteAllContents)

        End If
        File.Delete("c:\er\" & c.UserInfo.username & ".txt")
        frmWork = Nothing
        tCheckLogin.Stop()
        tCheckLogin.Dispose()
        tCheckUserInfo.Stop()
        tCheckUserInfo.Dispose()

        Me.Close()
        'ExitProcess(3838)
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '启动外部登录程序，然后以一秒为周期查询文件是否登录
        Try

            Me.Hide()
            If IsNothing(Me.frmWork) = False Then
                MsgBox("程序已在运行中。")
                Exit Sub
            End If
            'Dim s As String = InputBox("输入超级用户名")
            'If s = "luke" Then
            'frmWork = New Form1
            'frmWork.Show()
            'Me.Hide()
            'Exit Sub
            'End If
            Me.Visible = False
            If MakeDir() Then
                pId = Shell("c:\er\qy用户登录及余额查询_扣减软件.exe /" & dir & "/" & feeRate, 1)
                Me.Hide()
            Else
                MsgBox("创建用户名临时文件夹失败")
                Me.Close()
                Exit Sub
            End If

            hProcess = OpenProcess(&H1F0FFF, 0, pId)
            'Dim t As New Timer

            tCheckLogin.Interval = 1000
            tCheckLogin.Start()

            c = New Checker
            'Dim e As New FrmStart
            Me.Hide()
            'e.Show()
            Me.Visible = False
        Catch ex As Exception
            MsgBox(Err.Description)
        End Try
    End Sub

    Private Sub Form2_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Me.Hide()
    End Sub
End Class