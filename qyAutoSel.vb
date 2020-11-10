Public Class qySelectTool
    Structure typearr
        Dim arr() As String
    End Structure
    Function getzkbharr(ByVal rs As ADODB.Recordset) As String()
        Try
            Dim t() As String
            ReDim t(0)
            If rs.RecordCount = 0 Then
                Return t
            End If
            rs.MoveFirst()
            
            Do Until rs.EOF
                If t.Length = 1 And t(0) = "" Then t(0) = rs.Fields("zkbh").Value
                If t(t.Length - 1) <> rs.Fields("zkbh").Value Then
                    ReDim Preserve t(t.Length)
                    t(t.Length - 1) = rs.Fields("zkbh").Value
                End If
                rs.MoveNext()
            Loop
            Return t
        Catch ex As Exception
            MsgBox("getzkbharr过程出错-" & Err.Description)
        End Try
        
    End Function
    Function combinearr(ByVal typearr() As typearr) As String()

    End Function

End Class
