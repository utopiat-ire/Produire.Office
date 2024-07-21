' プロデル Microsoft Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Explicit On
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Namespace エクセル

	''' <summary>セル</summary>
	<種類(DocUrl:="/office/cell.htm")>
	Public Class セル
		Inherits DisposableObject
		Implements IProduireClass

		Dim RdrExcel As エクセル
		Public MyRange As Excel.Range

		Public Sub New(ByVal RdrExcel As エクセル, ByVal MyCell As Excel.Range)
			Me.RdrExcel = RdrExcel
			Me.MyRange = MyCell
			Use(MyCell)
		End Sub

#Region "手順"

		''' <summary>セルを選択状態にします</summary>
		''' <example>エクセルの現在シートの使用範囲を選択</example>
		<自分を>
		Public Sub 選択する()
			MyRange.Select()
		End Sub

		''' <summary>セルをコピーします</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub コピー()
			MyRange.Copy()
		End Sub

		''' <summary>セルを切り取ります</summary>
		'''　<remarks></remarks>
		<自分から, 動詞("カット", "切り取り")>
		Public Sub カット()
			MyRange.Cut()
		End Sub

		''' <summary>セルを貼り付けます</summary>
		'''　<remarks></remarks>
		<自分へ>
		Public Sub 貼り付け()
			MyRange.Select()
			MyRange.Worksheet.Paste()
		End Sub

		''' <summary>セルを削除します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 削除()
			MyRange.Delete()
		End Sub

		''' <summary>セルをコピーします</summary>
		'''　<remarks></remarks>
		<自分へ>
		Public Sub 挿入()
			MyRange.Insert()
		End Sub

		''' <summary>セルに含まれる値を置き換えます</summary>
		'''　<remarks>【置換前】を【置換後】へ</remarks>
		<自分で>
		Public Sub 置換(<を> ByVal 置換前 As String, <へ> ByVal 置換後 As String)
			MyRange.Replace(置換前, 置換後)
		End Sub

		''' <summary>セルを結合します</summary>
		<自分を>
		Public Sub 結合()
			MyRange.Merge()
		End Sub

		''' <summary>結合されたセルを分割します</summary>
		<自分を>
		Public Sub 分割()
			MyRange.UnMerge()
		End Sub

		''' <summary>セルに含まれるセルを並び替えます</summary>
		'''　<remarks>自分を【キー】で【方法】へ</remarks>
		<自分を, 動詞("並べ替える", "並べ替る", "並び替える", "並び替る", "並替える", "並替る")>
		Public Sub 並べ替える(<で> ByVal キー As セル, <へ, 省略> ByVal 方法 As String)
			Dim Order As Excel.XlSortOrder
			Dim Orient As Excel.XlSortOrientation
			If InStr(方法, "降順") > 0 Then Order = Excel.XlSortOrder.xlDescending
			If InStr(方法, "列") > 0 Then
				Orient = Excel.XlSortOrientation.xlSortRows
			Else
				Orient = Excel.XlSortOrientation.xlSortColumns
			End If
			MyRange.Sort(キー.MyRange, Order, Orientation:=Orient)
		End Sub

		''' <summary>セル範囲で指定した列で重複する行を削除します</summary>
		<自分から>
		Public Function 検索(<を> ByVal キーワード As String) As Excel.Range
			Return MyRange.Find(キーワード)
		End Function

		''' <summary>セルを検索します</summary>
		<自分で, 手順名("次を", "検索")>
		Public Function 次を検索(<から> ByVal セル As Excel.Range) As Excel.Range
			Return MyRange.FindNext(セル)
		End Function

		''' <summary>セルを検索します</summary>
		<自分で, 手順名("前を", "検索")>
		Public Function 前を検索(<から> ByVal セル As Excel.Range) As Excel.Range
			Return MyRange.FindNext(セル)
		End Function

		''' <summary>セル範囲で指定した列で重複する行を削除します</summary>
		<動詞("オートフィルタ")>
		Public Sub オートフィルタ(<自分序数詞("列", "を")> ByVal Column As Integer, <で> ByVal 条件 As String)
			MyRange.AutoFilter(Column, 条件)
		End Sub

		''' <summary>セル範囲で指定した列で重複する行を削除します</summary>
		<動詞("オートフィルタ")>
		Public Sub オートフィルタOr(<自分序数詞("列", "を")> ByVal Column As Integer, <助詞("と")> ByVal 条件1 As String, <で> ByVal 条件2 As String)
			MyRange.AutoFilter(Column, 条件1, Excel.XlAutoFilterOperator.xlOr, 条件2)
		End Sub

		''' <summary>セル範囲で指定した列で重複する行を削除します</summary>
		<自分から, 補語("重複を"), 動詞("削除")>
		Public Sub 重複を削除(<序数詞("列", "で")> ByVal Column As Integer)
			MyRange.RemoveDuplicates(Column, Excel.XlYesNoGuess.xlNo)
		End Sub

		''' <summary>セル範囲で指定した列で重複する行を削除します</summary>
		<自分から, 手順名("重複を", "削除")>
		Public Sub 重複を削除()
			MyRange.RemoveDuplicates(1, Excel.XlYesNoGuess.xlNo)
		End Sub

		''' <summary>セルを選択状態にします</summary>
		''' <example>エクセルの現在シートの使用範囲を選択</example>
		<自分へ>
		Public Function コメントを追加(<助詞("という")> 内容) As エクセルコメント
			Return New エクセルコメント(RdrExcel, MyRange.AddComment(内容))
		End Function

		''' <summary>セルを削除します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub クリア()
			MyRange.Clear()
		End Sub

		''' <summary>セルを削除します</summary>
		'''　<remarks></remarks>
		<自分から>
		Public Sub 値をクリア()
			MyRange.ClearContents()
		End Sub

		''' <summary>セルを削除します</summary>
		'''　<remarks></remarks>
		<自分から>
		Public Sub 書式をクリア()
			MyRange.ClearFormats()
		End Sub

		''' <summary>セルを削除します</summary>
		'''　<remarks></remarks>
		<自分から>
		Public Sub コメントをクリア()
			MyRange.ClearComments()
		End Sub

#End Region

#Region "設定項目"

		''' <summary>セルの内容</summary>
		Public Property 内容() As String
			Get
				Return MyRange.Text
			End Get
			Set(ByVal value As String)
				MyRange.Value = value
			End Set
		End Property

		''' <summary>セルの数式・関数式</summary>
		Public Property 計算式() As String
			Get
				Return MyRange.Formula
			End Get
			Set(ByVal Value As String)
				MyRange.Formula = Value
			End Set
		End Property

		''' <summary>セルの数式・関数式</summary>
		Public Property 計算式R1C1() As String
			Get
				Return MyRange.FormulaR1C1
			End Get
			Set(ByVal Value As String)
				MyRange.FormulaR1C1 = Value
			End Set
		End Property

		''' <summary>セルが結合されているかどうか</summary>
		Public Property 結合セル() As Boolean
			Get
				Return MyRange.MergeCells
			End Get
			Set(ByVal Value As Boolean)
				MyRange.MergeCells = Value
			End Set
		End Property

		''' <summary>セルの表示形式</summary>
		Public Property 表示形式() As String
			Get
				Return MyRange.NumberFormatLocal
			End Get
			Set(ByVal Value As String)
				MyRange.NumberFormatLocal = Value
			End Set
		End Property

		''' <summary>セルのフォント名</summary>
		Public Property フォント名() As String
			Get
				Return MyRange.Font.Name
			End Get
			Set(ByVal value As String)
				MyRange.Font.Name = value
			End Set
		End Property

		''' <summary>セルの文字サイズ</summary>
		<設定項目("文字サイズ", "フォントサイズ")>
		Public Property 文字サイズ() As Integer
			Get
				Return MyRange.Font.Size
			End Get
			Set(ByVal value As Integer)
				MyRange.Font.Size = value
			End Set
		End Property

		''' <summary>セルの背景色</summary>
		Public Property 背景色() As Color
			Get
				Return ColorTranslator.FromOle(MyRange.Interior.Color)
			End Get
			Set(ByVal value As Color)
				MyRange.Interior.Color = ColorTranslator.ToOle(value)
			End Set
		End Property

		''' <summary>セルの文字色</summary>
		Public Property 文字色() As Color
			Get
				Return ColorTranslator.FromOle(MyRange.Font.Color)
			End Get
			Set(ByVal value As Color)
				MyRange.Font.Color = ColorTranslator.ToOle(value)
			End Set
		End Property

		''' <summary>セルの文字を太字にするか</summary>
		Public Property 太字() As Boolean
			Get
				Return MyRange.Font.Bold
			End Get
			Set(ByVal value As Boolean)
				MyRange.Font.Bold = value
			End Set
		End Property

		''' <summary>セルの文字を斜体にするか</summary>
		Public Property 斜体() As Boolean
			Get
				Return MyRange.Font.Italic
			End Get
			Set(ByVal value As Boolean)
				MyRange.Font.Italic = value
			End Set
		End Property

		<列挙体(GetType(Excel.XlUnderlineStyle))>
		Public Enum 下線enum
			なし = Excel.XlUnderlineStyle.xlUnderlineStyleNone
			実線 = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
			二重線 = Excel.XlUnderlineStyle.xlUnderlineStyleDouble
		End Enum

		''' <summary>セルの文字に下線をするか</summary>
		Public Property 下線() As Excel.XlUnderlineStyle
			Get
				Return MyRange.Font.Underline
			End Get
			Set(ByVal value As Excel.XlUnderlineStyle)
				MyRange.Font.Underline = value
			End Set
		End Property

		''' <summary>セルの文字を取り消し線にするか</summary>
		Public Property 取り消し線() As Boolean
			Get
				Return MyRange.Font.Strikethrough
			End Get
			Set(ByVal value As Boolean)
				MyRange.Font.Strikethrough = value
			End Set
		End Property

		''' <summary>セルの文字に影を付けるか</summary>
		Public Property 影() As Boolean
			Get
				Return MyRange.Font.Shadow
			End Get
			Set(ByVal value As Boolean)
				MyRange.Font.Shadow = value
			End Set
		End Property

		''' <summary></summary>
		Public Property スタート() As String
			Get
				Return ""   'MyCell.Start
			End Get
			Set(ByVal value As String)

			End Set
		End Property

		''' <summary>セルの大きさ</summary>
		Public ReadOnly Property 大きさ() As SizeF
			Get
				Return New SizeF(MyRange.Width, MyRange.Height)
			End Get
		End Property

		''' <summary>セルの幅</summary>
		Public Property 幅() As Integer
			Get
				Return MyRange.Width
			End Get
			Set(ByVal value As Integer)
				MyRange.ColumnWidth = value
			End Set
		End Property

		''' <summary>セルの高さ</summary>
		Public Property 高さ() As Integer
			Get
				Return MyRange.Height
			End Get
			Set(ByVal value As Integer)
				MyRange.RowHeight = value
			End Set
		End Property

		''' <summary>セルの座標</summary>
		Public ReadOnly Property 座標() As PointF
			Get
				Return New PointF(MyRange.Left, MyRange.Top)
			End Get
		End Property

		''' <summary>セルの横位置</summary>
		Public ReadOnly Property 横() As Integer
			Get
				Return MyRange.Left
			End Get
		End Property

		''' <summary>セルの縦位置</summary>
		Public ReadOnly Property 縦() As Integer
			Get
				Return MyRange.Top
			End Get
		End Property

		''' <summary>選択しているセルの内容を配列形式で表します(セルを複数選択している場合)</summary>
		Public Property 一覧() As String()()
			Get
				Dim I As Integer, J As Integer
				Dim Arr()() As String, Arr2() As String

				With MyRange
					ReDim Arr(.Rows.Count - 1)
					For I = 0 To UBound(Arr)
						ReDim Arr2(.Columns.Count - 1)
						For J = 0 To UBound(Arr2)
							Dim Cell As Object = .Cells._Default(I + 1, J + 1)
							If TypeOf Cell Is String Then
								Arr2(J) = Cell
							Else
								Dim Cell2 As Excel.Range = Cell
								Arr2(J) = Cell2.Text
							End If
						Next
						Arr(I) = Arr2
					Next
				End With
				Return Arr
			End Get
			Set(ByVal value As String()())
				Dim I As Integer, J As Integer

				With MyRange
					For I = 0 To .Rows.Count
						For J = 0 To .Columns.Count
							.Cells._Default(I + 1, J + 1) = value(I)(J)
						Next
					Next
				End With
			End Set
		End Property

		''' <summary>選択しているセルの内容を配列形式で表します(セルを複数選択している場合)</summary>
		Public Property 値一覧() As String()()
			Get
				Dim I As Integer, J As Integer
				Dim Arr()() As String, Arr2() As String

				With MyRange
					ReDim Arr(.Rows.Count - 1)
					For I = 0 To UBound(Arr)
						ReDim Arr2(.Columns.Count - 1)
						For J = 0 To UBound(Arr2)
							Dim Cell As Object = .Cells._Default(I + 1, J + 1)
							If TypeOf Cell Is String Then
								Arr2(J) = Cell
							Else
								Dim Cell2 As Excel.Range = Cell
								Arr2(J) = Cell2.Value
							End If
						Next
						Arr(I) = Arr2
					Next
				End With
				Return Arr
			End Get
			Set(ByVal value As String()())
				一覧 = value
			End Set
		End Property
		''' <summary>選択セルの行数</summary>
		''' <returns>○</returns>
		Public ReadOnly Property 行数() As Integer
			Get
				Dim Cell As Excel.Range = MyRange
				Return Cell.Rows.Count
			End Get
		End Property

		''' <summary>選択セルの列数</summary>
		''' <returns>○</returns>
		Public ReadOnly Property 列数() As Integer
			Get
				Dim Cell As Excel.Range = MyRange
				Return Cell.Columns.Count
			End Get
		End Property

		''' <summary>選択セルの折り返し全体表示</summary>
		''' <returns>○</returns>
		Public Property 折り返し全体表示() As Boolean
			Get
				Return MyRange.WrapText
			End Get
			Set(value As Boolean)
				MyRange.WrapText = value
			End Set
		End Property
		''' <summary>選択セルの縮小全体表示</summary>
		''' <returns>○</returns>
		Public Property 縮小全体表示() As Boolean
			Get
				Return MyRange.ShrinkToFit
			End Get
			Set(value As Boolean)
				MyRange.ShrinkToFit = value
			End Set
		End Property

		<列挙体()>
		Public Enum 文字配置Enum
			左上
			左中央
			左下
			中央上
			中央
			中央下
			右上
			右中央
			右下
			標準
		End Enum

		''' <summary></summary>
		Public Property 文字配置() As 文字配置Enum
			Get

			End Get
			Set(ByVal value As 文字配置Enum)
				Dim HA, VA As Excel.Constants
				Select Case value
					Case 文字配置Enum.左上
						HA = Excel.Constants.xlLeft : VA = Excel.Constants.xlAbove
					Case 文字配置Enum.左中央
						HA = Excel.Constants.xlLeft : VA = Excel.Constants.xlCenter
					Case 文字配置Enum.左下
						HA = Excel.Constants.xlLeft : VA = Excel.Constants.xlBottom
					Case 文字配置Enum.中央上
						HA = Excel.Constants.xlCenter : VA = Excel.Constants.xlAbove
					Case 文字配置Enum.中央
						HA = Excel.Constants.xlCenter : VA = Excel.Constants.xlCenter
					Case 文字配置Enum.中央下
						HA = Excel.Constants.xlCenter : VA = Excel.Constants.xlBottom
					Case 文字配置Enum.右上
						HA = Excel.Constants.xlRight : VA = Excel.Constants.xlAbove
					Case 文字配置Enum.右中央
						HA = Excel.Constants.xlRight : VA = Excel.Constants.xlCenter
					Case 文字配置Enum.右下
						HA = Excel.Constants.xlRight : VA = Excel.Constants.xlBottom
					Case 文字配置Enum.標準
						HA = Excel.Constants.xlNone : VA = Excel.Constants.xlNone
					Case Else : Throw New ProduireException(value & "という設定値はありません。", エクセル.ERRORBASE + 9)
				End Select
				MyRange.HorizontalAlignment = HA
				MyRange.VerticalAlignment = VA
			End Set
		End Property

		<列挙体()>
		Public Enum 形Enum
			四角
			上
			下
			右
			左
			すべて
			中枠
			なし
		End Enum

		''' <summary></summary>
		Public Property 罫線の形() As 形Enum
			Get

			End Get
			Set(ByVal value As 形Enum)
				Dim I As Integer
				Dim CellFrame(5) As Boolean
				Dim AllClear As Boolean
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				For I = 0 To UBound(CellBorders)
					If Not CellBorders(I) Is Nothing Then CellFrame(I) = (CellBorders(I).LineStyle = Excel.XlLineStyle.xlContinuous)
				Next
				Select Case value
					Case 形Enum.四角
						CellFrame(0) = True : CellFrame(1) = True : CellFrame(2) = True : CellFrame(3) = True
					Case 形Enum.上
						CellFrame(0) = True
					Case 形Enum.下
						CellFrame(2) = True
					Case 形Enum.右
						CellFrame(1) = True
					Case 形Enum.左
						CellFrame(3) = True
					Case 形Enum.すべて
						CellFrame(0) = True : CellFrame(1) = True : CellFrame(2) = True : CellFrame(3) = True : CellFrame(4) = True : CellFrame(5) = True
					Case 形Enum.中枠
						CellFrame(0) = False : CellFrame(1) = False : CellFrame(2) = False : CellFrame(3) = False : CellFrame(4) = True : CellFrame(5) = True
					Case 形Enum.なし
						CellFrame(0) = False : CellFrame(1) = False : CellFrame(2) = False : CellFrame(3) = False : CellFrame(4) = False : CellFrame(5) = False
						AllClear = True
					Case Else : Throw New ProduireException(value & "という設定値はありません。", エクセル.ERRORBASE + 9)
				End Select

				For I = 0 To UBound(CellBorders)
					If Not CellBorders(I) Is Nothing Then
						With CellBorders(I)
							If CellFrame(I) Or AllClear Then
								.LineStyle = IIf(CellFrame(I), Excel.XlLineStyle.xlContinuous, Excel.XlLineStyle.xlLineStyleNone)
							End If
						End With
					End If
				Next
			End Set
		End Property

		<列挙体(GetType(Excel.XlLineStyle))>
		Public Enum スタイルEnum
			なし = Excel.XlLineStyle.xlLineStyleNone
			実線 = Excel.XlLineStyle.xlContinuous
			破線 = Excel.XlLineStyle.xlDash
			一点鎖線 = Excel.XlLineStyle.xlDashDot
			二点鎖線 = Excel.XlLineStyle.xlDashDotDot
			点線 = Excel.XlLineStyle.xlDot
			二重線 = Excel.XlLineStyle.xlDouble
			斜線 = Excel.XlLineStyle.xlSlantDashDot
		End Enum

		''' <summary></summary>
		Public Property 罫線のスタイル() As Excel.XlLineStyle
			Get
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				Return CellBorders(0).LineStyle
			End Get
			Set(ByVal value As Excel.XlLineStyle)
				Dim I As Integer
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				For I = 0 To UBound(CellBorders)
					If Not CellBorders(I) Is Nothing Then
						With CellBorders(I)
							If .LineStyle = Excel.XlLineStyle.xlContinuous Then
								.LineStyle = value
							End If
						End With
					End If
				Next
			End Set
		End Property

		<列挙体(GetType(Excel.XlBorderWeight))>
		Public Enum 太さEnum
			極細 = Excel.XlBorderWeight.xlHairline
			細い = Excel.XlBorderWeight.xlThin
			中 = Excel.XlBorderWeight.xlMedium
			太い = Excel.XlBorderWeight.xlThick
		End Enum

		''' <summary></summary>
		Public Property 罫線の太さ() As Excel.XlBorderWeight
			Get
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				Return CellBorders(0).Weight
			End Get
			Set(ByVal value As Excel.XlBorderWeight)
				Dim I As Integer
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				For I = 0 To UBound(CellBorders)
					If Not CellBorders(I) Is Nothing Then
						With CellBorders(I)
							If .LineStyle = Excel.XlLineStyle.xlContinuous Then
								.Weight = value
							End If
						End With
					End If
				Next
			End Set
		End Property

		''' <summary></summary>
		Public Property 罫線の色() As Color
			Get
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				Return ColorTranslator.FromOle(CellBorders(0).Color)
			End Get
			Set(ByVal value As Color)
				Dim I As Integer
				Dim CellBorders() As Excel.Border
				CellBorders = RdrExcel.pCellBorders(MyRange)
				For I = 0 To UBound(CellBorders)
					If Not CellBorders(I) Is Nothing Then
						With CellBorders(I)
							If .LineStyle = Excel.XlLineStyle.xlContinuous Then
								.Color = ColorTranslator.ToOle(value)
							End If
						End With
					End If
				Next
			End Set
		End Property
		<列挙体(GetType(Excel.XlLineStyle))>
		Public Enum 枠Enum
			なし = Excel.XlLineStyle.xlLineStyleNone
			実線 = Excel.XlLineStyle.xlContinuous
			破線 = Excel.XlLineStyle.xlDash
			一点鎖線 = Excel.XlLineStyle.xlDashDot
			二点鎖線 = Excel.XlLineStyle.xlDashDotDot
			点線 = Excel.XlLineStyle.xlDot
			二重線 = Excel.XlLineStyle.xlDouble
			斜線 = Excel.XlLineStyle.xlSlantDashDot
		End Enum

		''' <summary></summary>
		Public Property 枠() As Excel.XlLineStyle
			Get
				Return CType(MyRange.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle, Excel.XlLineStyle)
			End Get
			Set(ByVal value As Excel.XlLineStyle)
				With MyRange
					.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = value
					.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = value
					.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = value
					.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = value
				End With
			End Set
		End Property

		''' <summary></summary>
		Public Property セルの罫線() As Boolean
			Get
				Return CType(MyRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle, Excel.XlLineStyle) <> Excel.XlLineStyle.xlLineStyleNone
			End Get
			Set(ByVal value As Boolean)
				Dim F As Excel.XlLineStyle
				With MyRange
					'項目欄を二重線で区切る
					'UPGRADE_WARNING: オブジェクト Ob.Borders.LineStyle の既定プロパティを解決できませんでした。 詳細については、'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"' をクリックしてください。
					If value Then F = CType(.Borders.LineStyle, Excel.XlLineStyle) Else F = Excel.XlLineStyle.xlLineStyleNone
					.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = F
					.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = F
				End With
			End Set
		End Property

		''' <summary></summary>
		Public ReadOnly Property コメント() As エクセルコメント
			Get
				Return New エクセルコメント(RdrExcel, MyRange.Comment)
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Excel.Range
			Get
				Return MyRange
			End Get
		End Property

#End Region

	End Class

	''' <summary>エクセル図形</summary>
	<種類(DocUrl:="/office/excelshape.htm")>
	Public Class エクセル図形
		Inherits DisposableObject
		Implements IProduireClass

		ReadOnly ExcelApp As Excel.Application
		Dim MyShape As Excel.Shape
		ReadOnly RdrExcel As エクセル

		Public Sub New(ByVal RdrExcel As エクセル, ByVal MyShape As Excel.Shape)
			Me.RdrExcel = RdrExcel
			Me.MyShape = MyShape
			Me.ExcelApp = MyShape.Application
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
				Return MyShape.Visible
			End Get
			Set(ByVal Value As Boolean)
				MyShape.Visible = Value
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
					Return New Rectangle(.Left, .Top, .Width, .Height)
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

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Excel.Shape
			Get
				Return MyShape
			End Get
		End Property

#End Region

	End Class


	''' <summary>エクセル図形</summary>
	<種類(DocUrl:="/office/excelshape.htm")>
	Public Class エクセルコメント
		Inherits DisposableObject
		Implements IProduireClass

		ReadOnly ExcelApp As Excel.Application
		Dim MyComment As Excel.Comment
		ReadOnly RdrExcel As エクセル

		Public Sub New(ByVal RdrExcel As エクセル, ByVal MyComment As Excel.Comment)
			Me.RdrExcel = RdrExcel
			Me.MyComment = MyComment
			Me.ExcelApp = MyComment.Application
			Use(MyComment)
		End Sub

#Region "手順"

		''' <summary>図形を削除します</summary>
		<自分を>
		Public Sub 消す()
			MyComment.Delete()
		End Sub

#End Region

#Region "設定項目"
		''' <summary>図形の名前</summary>
		Public Property 内容() As String
			Get
				Return MyComment.Text()
			End Get
			Set(ByVal Value As String)
				MyComment.Text(Value)
			End Set
		End Property

		''' <summary>図形を表示するかどうか</summary>
		Public Property 表示() As Boolean
			Get
				Return MyComment.Visible
			End Get
			Set(ByVal Value As Boolean)
				MyComment.Visible = Value
			End Set
		End Property

		''' <summary>図形の幅</summary>
		Public ReadOnly Property 図形() As エクセル図形
			Get
				Return New エクセル図形(RdrExcel, MyComment.Shape)
			End Get
		End Property

		''' <summary>作成者</summary>
		Public ReadOnly Property 作成者() As Integer
			Get
				Return MyComment.Author
			End Get
		End Property

		''' <summary>生成アプリ</summary>
		Public ReadOnly Property 生成アプリ() As Integer
			Get
				Return MyComment.Creator
			End Get
		End Property

		''' <summary>次</summary>
		Public ReadOnly Property 次() As エクセルコメント
			Get
				Return New エクセルコメント(RdrExcel, MyComment.Next())
			End Get
		End Property

		''' <summary>前</summary>
		Public ReadOnly Property 前() As エクセルコメント
			Get
				Return New エクセルコメント(RdrExcel, MyComment.Previous())
			End Get
		End Property

		''' <summary>コメントの親</summary>
		Public ReadOnly Property 親() As セル
			Get
				Return New セル(RdrExcel, MyComment.Parent)
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Excel.Comment
			Get
				Return MyComment
			End Get
		End Property

#End Region

	End Class
End Namespace
