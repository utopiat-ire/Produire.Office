' プロデル Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Strict On
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Namespace ワード

	''' <summary>ワード選択範囲</summary>
	<種類(DocUrl:="/office/wordselection.htm")>
	Public Class ワード選択範囲
		Inherits DisposableObject
		Implements IProduireClass

		ReadOnly WordApp As Word.Application
		Dim MySelect As Word.Selection
		ReadOnly RdrWord As ワード

		Public Sub New(ByVal RdrWord As ワード, ByVal MySelect As Word.Selection)
			Me.RdrWord = RdrWord
			Me.MySelect = MySelect
			Use(MySelect)
		End Sub

#Region "手順"

		''' <summary>選択している文字をクリップボードへコピーします</summary>
		<自分を>
		Public Sub コピー()
			On Error Resume Next
			MySelect.Copy()
			If Err.Number <> 0 Then
				Throw New ProduireException(Err.Description, Err.Number)
			End If
		End Sub

		''' <summary>選択している文字をクリップボードへコピーして、切り取ります</summary>
		<自分から, 動詞("カット", "切り取り")>
		Public Sub カット()
			MySelect.Cut()
		End Sub

		''' <summary>クリップボードにある文字を貼り付けます</summary>
		<自分へ>
		Public Sub 貼り付け()
			MySelect.Paste()
		End Sub

		''' <summary>選択している文字を消します</summary>
		<自分を>
		Public Sub 消す()
			MySelect.Delete()
		End Sub

		''' <summary>選択している文字を消します</summary>
		<自分で, 補語("書式を")>
		Public Sub 書式をクリアする()
			MySelect.ClearFormatting()
		End Sub

		''' <summary>選択範囲からキーワードを探します</summary>
		<自分から>
		Public Function 検索する() As Boolean
			Dim Ret As Boolean
			With MySelect.Find
				Ret = .Execute()
			End With
			Return Ret
		End Function

		''' <summary>選択範囲からキーワードを探します</summary>
		<自分から>
		Public Function 検索する(<を> キーワード As String) As Boolean
			Dim Ret As Boolean
			With MySelect.Find
				.ClearFormatting()
				.Text = キーワード
				Ret = .Execute()
			End With
			Return Ret
		End Function

		''' <summary>選択範囲からキーワードを探します</summary>
		<自分で>
		Public Function 置換する(<を> キーワード As String, <へ> 置換後 As String) As Boolean
			Dim Ret As Boolean
			With MySelect.Find
				.ClearFormatting()
				.Text = キーワード
				.Replacement.ClearFormatting()
				.Replacement.Text = 置換後
				.Execute(Replace:=Word.WdReplace.wdReplaceOne)
			End With
			Return Ret
		End Function

		''' <summary>選択範囲からキーワードを探します</summary>
		<自分で>
		Public Function すべて置換する(<を> キーワード As String, <へ> 置換後 As String) As Boolean
			Dim Ret As Boolean
			With MySelect.Find
				.ClearFormatting()
				.Text = キーワード
				.Replacement.ClearFormatting()
				.Replacement.Text = 置換後
				.Execute(Replace:=Word.WdReplace.wdReplaceAll)
			End With
			Return Ret
		End Function

#End Region

#Region "設定項目"

		''' <summary>フォント名を指定します</summary>
		Public Property フォント名() As String
			Get
				Return MySelect.Font.Name
			End Get
			Set(ByVal value As String)
				MySelect.Font.Name = value
			End Set
		End Property

		''' <summary>文字サイズ</summary>
		<設定項目("文字サイズ", "フォントサイズ")>
		Public Property 文字サイズ() As Single
			Get
				Return MySelect.Font.Size
			End Get
			Set(ByVal value As Single)
				MySelect.Font.Size = value
			End Set
		End Property

		''' <summary>文字の色</summary>
		Public Property 文字色() As Color
			Get
				Return ColorTranslator.FromOle(MySelect.Font.Color)
			End Get
			Set(ByVal value As Color)
				MySelect.Font.Color = CType(ColorTranslator.ToOle(value), Word.WdColor)
			End Set
		End Property

		''' <summary>文字を太字にするかどうか</summary>
		Public Property 太字() As Boolean
			Get
				Return CBool(MySelect.Font.Bold)
			End Get
			Set(ByVal value As Boolean)
				MySelect.Font.Bold = CInt(value)
			End Set
		End Property

		''' <summary>文字を斜体にするかどうか</summary>
		Public Property 斜体() As Boolean
			Get
				Return CBool(MySelect.Font.Italic)
			End Get
			Set(ByVal value As Boolean)
				MySelect.Font.Italic = CInt(value)
			End Set
		End Property

		''' <summary>文字を下線を引くかどうか</summary>
		Public Property 下線() As Boolean
			Get
				Return CBool(MySelect.Font.Underline)
			End Get
			Set(ByVal value As Boolean)
				MySelect.Font.Underline = CType(value, Word.WdUnderline)
			End Set
		End Property

		''' <summary>選択されている部分の内容を表します
		''' 文字が選択されていない場合は、カーソルがある位置へ文字を挿入します</summary>
		Public Property 内容() As String
			Get
				Return MySelect.Text
			End Get
			Set(ByVal value As String)
				MySelect.Text = value
			End Set
		End Property

		''' <summary>選択する先頭の場所を文字数で指定します</summary>
		Public Property 先頭位置() As Integer
			Get
				Return MySelect.Start
			End Get
			Set(ByVal value As Integer)
				MySelect.SetRange(value, 0)
			End Set
		End Property

		''' <summary>選択する文字数を指定します</summary>
		Public Property 長さ() As Integer
			Get
				With MySelect
					Return .End - .Start
				End With
			End Get
			Set(ByVal value As Integer)
				With MySelect
					.SetRange(.Start, value)
				End With
			End Set
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Word.Selection
			Get
				Return MySelect
			End Get
		End Property

#End Region

		Protected Overrides Sub Finalize()
			Marshal.FinalReleaseComObject(MySelect)
			MyBase.Finalize()
		End Sub
	End Class

	''' <summary>ワード図形</summary>
	<種類(DocUrl:="/office/wordshape.htm")>
	Public Class ワード図形
		Inherits DisposableObject
		Implements IProduireClass

		ReadOnly WordApp As Word.Application
		ReadOnly RdrWord As ワード
		Dim MyShape As Word.Shape

		Public Sub New(ByVal RdrWord As ワード, ByVal MyShape As Word.Shape)
			Me.RdrWord = RdrWord
			Me.MyShape = MyShape
			Use(MyShape)
		End Sub

#Region "手順"

		''' <summary>図形を選択します</summary>
		<自分を>
		Public Sub 選択()
			MyShape.Select()
		End Sub

		''' <summary>図形を削除します</summary>
		<自分を>
		Public Sub 消す()
			MyShape.Delete()
		End Sub

		''' <summary>図形を[コピー先]の図形の後ろへコピーします</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub コピー()
			MyShape.Duplicate()
		End Sub

#End Region

#Region "設定項目"

		''' <summary>図形の名前</summary>
		Public Property 名前() As String
			Get
				Return MyShape.Name
			End Get
			Set(ByVal Value As String)
				MyShape.Name = Value
			End Set
		End Property

		''' <summary>図形を表示するかどうか</summary>
		Public Property 表示() As Boolean
			Get
				Return MyShape.Visible = Microsoft.Office.Core.MsoTriState.msoCTrue
			End Get
			Set(ByVal Value As Boolean)
				MyShape.Visible = CType(IIf(Value, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoFalse), Microsoft.Office.Core.MsoTriState)
			End Set
		End Property

		''' <summary>図形の幅</summary>
		Public Property 幅() As Single
			Get
				Return MyShape.Width
			End Get
			Set(ByVal Value As Single)
				MyShape.Width = Value
			End Set
		End Property

		''' <summary>図形の高さ</summary>
		Public Property 高さ() As Single
			Get
				Return MyShape.Height
			End Get
			Set(ByVal Value As Single)
				MyShape.Height = Value
			End Set
		End Property

		''' <summary>横の位置</summary>
		Public Property 横() As Single
			Get
				Return MyShape.Left
			End Get
			Set(ByVal Value As Single)
				MyShape.Left = Value
			End Set
		End Property

		''' <summary>縦の位置</summary>
		Public Property 縦() As Single
			Get
				Return MyShape.Top
			End Get
			Set(ByVal Value As Single)
				MyShape.Top = Value
			End Set
		End Property

		''' <summary>図形の位置</summary>
		Public Property 位置() As PointF
			Get
				Return New PointF(MyShape.Left, MyShape.Top)
			End Get
			Set(ByVal Value As PointF)
				MyShape.Left = Value.X
				MyShape.Top = Value.Y
			End Set
		End Property

		''' <summary>図形の大きさ</summary>
		Public Property 大きさ() As SizeF
			Get
				Return New SizeF(MyShape.Width, MyShape.Height)
			End Get
			Set(ByVal Value As SizeF)
				MyShape.Width = Value.Width
				MyShape.Height = Value.Height
			End Set
		End Property

		''' <summary>図形の位置と大きさ</summary>
		Public Property 位置と大きさ() As RectangleF
			Get
				With MyShape
					Return New RectangleF(.Left, .Top, .Width, .Height)
				End With
			End Get
			Set(ByVal Value As RectangleF)
				With MyShape
					.Left = Value.X
					.Top = Value.Y
					.Width = Value.Width
					.Height = Value.Height
				End With
			End Set
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Word.Shape
			Get
				Return MyShape
			End Get
		End Property

#End Region

		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub

	End Class

End Namespace
