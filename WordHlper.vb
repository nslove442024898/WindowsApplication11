Imports System.IO

Imports MSWord = Microsoft.Office.Interop.Word


Module WordHelper

    Public Sub SetTextBoxTextFont(ByRef wordApp As MSWord.Application, sampleWordFullName As String, workDir As String, fontnames() As String, rowDistance As Integer, tcbhxx As MyTCFill2WordData)

        Dim curDocName = Path.Combine(workDir, $"{tcbhxx.ZKBH}-钻探记录.doc")

        Dim curDocObj = wordApp.Documents.Open(sampleWordFullName)

        Dim FontSize() As Integer = {10.5, 6.5}
        '指定五种字号

        'Dim FontName(0 to 1)
        ''字体名称在2种字体之间进行波动，可改写，但需要保证系统拥有下列字体
        'FontName(1) = "liguofu"
        'FontName(0) = "【嵐】芊柔体"

        Dim textBox As MSWord.Shape
        '填写表头
        Dim tchs As Integer = tcbhxx.ListTCinfor.Count

        Dim tcxxs01(0 To 6) As MSWord.Shape

        'Dim tcxxsAll As New List(Of MSWord.Shape()) '从第二行开始的数据

        For index = 1 To curDocObj.Shapes.Count Step 1
            textBox = curDocObj.Shapes(index)
            If textBox.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox And textBox.AutoShapeType = Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle Then
                'msoTextBox===>17,MsoShapeType===TextBox.Type.ToString()
                With textBox.TextFrame
                    .TextRange.Font.Name = fontnames(0)
                    .MarginBottom = 0
                    .MarginLeft = 0
                    .MarginRight = 0
                    .MarginTop = 0
                End With
                'oCtl.TextFrame.TextRange.Text.Replace("Invoice", "Invoice - Paid: " + paid_date)
                Select Case textBox.TextFrame.TextRange.Text.TrimEnd()
                    Case "GCMC"
                        textBox.TextFrame.TextRange.Text = tcbhxx.GCMC
                    Case "ZKBH"
                        textBox.TextFrame.TextRange.Text = tcbhxx.ZKBH
                    Case "ZKBG"
                        textBox.TextFrame.TextRange.Text = Math.Round(CDbl(tcbhxx.ZKBG), 3)
                    Case "ZKZJLX"
                        textBox.TextFrame.TextRange.Text = tcbhxx.ZKZJLX
                    Case "KCDW"
                        textBox.TextFrame.TextRange.Text = tcbhxx.KCDW
                    Case "YR"
                        textBox.Top = textBox.Top + 6
                        textBox.Left = textBox.Left - 2
                        textBox.TextFrame.TextRange.Text = tcbhxx.YR
                    Case "MO"
                        textBox.Top = textBox.Top + 6
                        textBox.Left = textBox.Left + 10
                        textBox.TextFrame.TextRange.Text = tcbhxx.MO
                    Case "DA"
                        textBox.Top = textBox.Top + 6
                        textBox.Left = textBox.Left + 10
                        textBox.TextFrame.TextRange.Text = tcbhxx.DA
                    Case "YR1"
                        textBox.Top = textBox.Top + 5
                        textBox.Left = textBox.Left - 2

                        textBox.TextFrame.TextRange.Text = tcbhxx.YR1
                    Case "MO1"
                        textBox.Left = textBox.Left + 10
                        textBox.Top = textBox.Top + 4
                        textBox.TextFrame.TextRange.Text = tcbhxx.MO1
                    Case "DA1"
                        textBox.Left = textBox.Left + 6
                        textBox.Top = textBox.Top + 4
                        textBox.TextFrame.TextRange.Text = tcbhxx.DA1
                    Case "TCCDSD"
                        tcxxs01(0) = textBox
                    Case "TCMC"
                        tcxxs01(1) = textBox

                    Case "YS"
                        tcxxs01(2) = textBox

                    Case "KSX"
                        tcxxs01(3) = textBox

                    Case "MSD"
                        tcxxs01(4) = textBox

                    Case "SD"
                        tcxxs01(5) = textBox

                    Case "MS"
                        tcxxs01(6) = textBox
                End Select

            End If
        Next index

        Dim tcxxsindex(0 To 6) As MSWord.Shape

        'Dim MyTCFill2WordDataTCinfoType As Type = GetType(MyTCFill2WordDataTCinfo)

        For index = 1 To tchs - 1
            For Each item As MSWord.Shape In tcxxs01
                Select Case item.TextFrame.TextRange.Text.Trim()
                    Case "TCCDSD"
                        tcxxsindex（0） = item.Duplicate()
                        tcxxsindex（0).Top = item.Top + index * rowDistance
                        tcxxsindex（0).Left = item.Left
                        tcxxsindex（0).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).TCCDSD
                        tcxxsindex（0).TextFrame.TextRange.Font.Bold = False
                    Case "TCMC"
                        tcxxsindex（1） = item.Duplicate()
                        tcxxsindex（1).Top = item.Top + index * rowDistance
                        tcxxsindex（1).Left = item.Left
                        tcxxsindex（1).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).TCMC
                        tcxxsindex（1).TextFrame.TextRange.Font.Bold = False
                    Case "YS"
                        tcxxsindex（2） = item.Duplicate()
                        tcxxsindex（2).Top = item.Top + index * rowDistance
                        tcxxsindex（2).Left = item.Left
                        tcxxsindex（2).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).YS
                        tcxxsindex（2).TextFrame.TextRange.Font.Bold = False
                    Case "KSX"
                        tcxxsindex（3） = item.Duplicate()
                        tcxxsindex（3).Top = item.Top + index * rowDistance
                        tcxxsindex（3).Left = item.Left
                        tcxxsindex（3).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).KSX
                        tcxxsindex（3).TextFrame.TextRange.Font.Bold = False
                    Case "MSD"
                        tcxxsindex（4） = item.Duplicate()
                        tcxxsindex（4).Top = item.Top + index * rowDistance
                        tcxxsindex（4).Left = item.Left
                        tcxxsindex（4).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).MSD
                        tcxxsindex（4).TextFrame.TextRange.Font.Bold = False
                    Case "SD"
                        tcxxsindex（5） = item.Duplicate()
                        tcxxsindex（5).Top = item.Top + index * rowDistance
                        tcxxsindex（5).Left = item.Left
                        tcxxsindex（5).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).SD
                        tcxxsindex（5).TextFrame.TextRange.Font.Bold = False
                    Case "MS"
                        tcxxsindex（6） = item.Duplicate()
                        tcxxsindex（6).Top = item.Top + index * rowDistance
                        tcxxsindex（6).Left = item.Left
                        tcxxsindex（6).TextFrame.TextRange.Text = tcbhxx.ListTCinfor(index).MS
                        tcxxsindex（6).TextFrame.TextRange.Font.Bold = False
                End Select
            Next
        Next


        For Each item As MSWord.Shape In tcxxs01
            Select Case item.TextFrame.TextRange.Text.Trim()
                Case "TCCDSD"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).TCCDSD
                Case "TCMC"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).TCMC
                Case "YS"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).YS
                Case "KSX"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).KSX
                Case "MSD"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).MSD
                Case "SD"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).SD
                Case "MS"
                    item.TextFrame.TextRange.Text = tcbhxx.ListTCinfor(0).MS
            End Select
            item.TextFrame.TextRange.Bold = False
        Next

        'curWordDoc.GetType().InvokeMember("Close", Reflection.BindingFlags.InvokeMethod, curWordDoc, Nothing, {True})
        curDocObj.SaveAs(curDocName)
        wordApp.DisplayAlerts = MSWord.WdAlertLevel.wdAlertsNone
        curDocObj.Close(False)

    End Sub
End Module
