' プロデル Microsoft Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Strict On
Imports System.Drawing
Imports System.Collections.Generic
Imports Microsoft.Office.Interop
Imports MsoTriState = Microsoft.Office.Core.MsoTriState

Namespace パワーポイント

	''' <summary>パワーポイント図形</summary>
	<種類(DocUrl:="/office/powershape.htm")>
	Public Class パワーポイント図形
		Inherits DisposableObject
		Implements IProduireClass

		ReadOnly PptApp As PowerPoint.Application
		Dim MyShape As PowerPoint.Shape
		ReadOnly RdrPpt As パワーポイント

		Public Sub New(ByVal RdrPpt As パワーポイント, ByVal MyShape As PowerPoint.Shape)
			Me.RdrPpt = RdrPpt
			Me.MyShape = MyShape
			Me.PptApp = CType(MyShape.Application, PowerPoint.Application)
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
			MyShape.Copy()
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
				Return MyShape.Visible = MsoTriState.msoCTrue Or MyShape.Visible = MsoTriState.msoTrue
			End Get
			Set(ByVal Value As Boolean)
				If Value Then
					MyShape.Visible = MsoTriState.msoTrue
				Else
					MyShape.Visible = MsoTriState.msoFalse
				End If
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
		Public Property 位置と大きさ() As Rectangle
			Get
				With MyShape
					Return New Rectangle(CInt(.Left), CInt(.Top), CInt(.Width), CInt(.Height))
				End With
			End Get
			Set(ByVal Value As Rectangle)
				With MyShape
					.Left = Value.X
					.Top = Value.Y
					.Width = Value.Width
					.Height = Value.Height
				End With
			End Set
		End Property

#End Region

	End Class

End Namespace
