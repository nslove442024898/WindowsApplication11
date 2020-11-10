Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

Namespace GetMyDGVS
	Public Module [Myclass]
		Public Function GetControlsofType(Of T As Control)(root As Control) As IList(Of T)
			Dim list As List(Of T) = New List(Of T)()
			For Each control As Control In root.Controls
                Dim t1 As T = TryCast(control, T)
                Dim flag As Boolean = t1 IsNot Nothing
                If flag Then
                    list.Add(t1)
                End If
				list.AddRange([Myclass].GetControlsofType(Of T)(control))
			Next
			Return list
		End Function

		Public Function GetControlsofFrm(Of T As Control)(frm As Form) As IList(Of T)
			Dim list As List(Of T) = New List(Of T)()
			For Each control As Control In frm.Controls
				Dim flag As Boolean = TypeOf control Is T
				If flag Then
					list.Add(TryCast(control, T))
				End If
				list.AddRange([Myclass].GetControlsofType(Of T)(control))
			Next
			Return list
		End Function
	End Module
End Namespace
