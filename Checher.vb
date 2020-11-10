Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports System.Text
Imports ADODB
Imports System.IO
Imports System.Security.Cryptography
Public Structure userInfo
    Dim username As String
    Dim userID As UserIdentity
    Dim freeCount As Integer
    Dim money As Double

    Dim validDate As String

End Structure
Public Enum UserIdentity
    Administrator = 1
    TrialUser = 2
    User = 3
    UserNotExist = 4
End Enum



Public Class Checker


    Public UserInfo As userInfo
    Public Event RaiseAlert()
    Public Event UserTimeOut()
    Public Event UserMoneyOut()
    Public userNameDocPath As String



    'Public rs As ADODB.Recordset
    Public Sub New()

    End Sub





    Sub ReadUserInfo()
        Try

            Dim str() As String = File.ReadAllLines(userNameDocPath)
            Dim s() As String = File.ReadAllLines("c:\er\" & str(0) & "\checker.txt")
            Try
                With userInfo
                    .username = s(0)
                    If s(1) = "用户类型：体验用户" Then .userID = UserIdentity.TrialUser
                    If s(1) = "用户类型：正式用户" Then .userID = UserIdentity.User
                    If s(1) = "用户类型：管理员" Then .userID = UserIdentity.Administrator
                    If .userID = UserIdentity.TrialUser Then .validDate = s(2)
                    If .userID = UserIdentity.User Then .money = s(2)

                End With
                Select Case UserInfo.userID
                    Case UserIdentity.TrialUser

                        If DateDiff(DateInterval.Minute, Now, CDate(UserInfo.validDate)) < 5 And DateDiff(DateInterval.Minute, Now, CDate(UserInfo.validDate)) > 0 Then
                            RaiseEvent RaiseAlert()
                        ElseIf DateDiff(DateInterval.Minute, Now, CDate(UserInfo.validDate)) <= 0 Then
                            RaiseEvent UserTimeOut()
                        End If
                    Case UserIdentity.User

                        If UserInfo.money < 5 And UserInfo.money > 0 Then
                            RaiseEvent RaiseAlert()
                        ElseIf UserInfo.money <= 0 Then
                            RaiseEvent UserMoneyOut()

                        End If
                End Select
            Catch ex As Exception

            End Try
        Catch ex As Exception
            MsgBox("读取用户信息过程出错" & Err.Description)
        End Try
    End Sub


End Class
