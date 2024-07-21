' プロデル Microsoft Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Strict On
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Module Tips
	Friend Sub FinishComObject(ByRef ComObject As Object)
		'COM オブジェクトの使用後、明示的に COM オブジェクトへの参照を解放する
		Try
			'提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
			If Not ComObject Is Nothing AndAlso Marshal.IsComObject(ComObject) Then
				Dim I As Integer
				Do
					I = Marshal.ReleaseComObject(ComObject)
				Loop Until I <= 0
			End If
		Catch
		Finally
			'参照を解除する
			ComObject = Nothing
		End Try
	End Sub

	<DllImport("user32.dll")>
	Friend Function SetForegroundWindow(hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
	End Function

End Module

Public Class DisposableObject
	Implements IDisposable
	Protected DisposeList As New List(Of Object)

	Private disposedValue As Boolean = False		' 重複する呼び出しを検出するため

	' IDisposable
	Protected Overridable Sub Dispose(ByVal disposing As Boolean)
		On Error Resume Next
		If Not Me.disposedValue Then
			If disposing Then
				Dim Item As Object
				For Each Item In DisposeList
					If TypeOf Item Is IDisposable Then
						CType(Item, IDisposable).Dispose()
					Else
						Marshal.FinalReleaseComObject(Item)
					End If
				Next
				DisposeList.Clear()
			End If
		End If
		Me.disposedValue = True
	End Sub

#Region " IDisposable Support "
	' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
	Public Sub Dispose() Implements IDisposable.Dispose
		' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub
#End Region

	''' <summary>
	''' 手動で破棄するCOMオブジェクト/IDisposableとしてリストへ追加します
	''' </summary>
	''' <param name="Item"></param>
	''' <remarks></remarks>
	Protected Sub Use(ByVal Item As Object)
		DisposeList.Add(Item)
	End Sub

	Protected Overrides Sub Finalize()
		Dispose()
	End Sub
End Class