' プロデル Microsoft Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Strict On
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports utopiat

Namespace エクセル

	''' <summary></summary>
	<種類(DocUrl:="/office/excel.htm")>
	Public Class エクセル
		Inherits DisposableObject
		Implements IProduireStaticClass
		Implements IObjectContainer

		Dim ExcelApp As Excel.Application
		Dim BookList As WorkbookCollection
		Public Const ERRORBASE As Integer = 2000

		Public Shared StaticExcelApp As Excel.Application

		Friend Function pGetWorkSheet(ByRef SheetName As String) As Excel.Worksheet
			If IsNumeric(SheetName) Then
				pGetWorkSheet = CType(ExcelApp.Worksheets(CShort(SheetName)), Excel.Worksheet)
			Else
				pGetWorkSheet = CType(ExcelApp.Worksheets(SheetName), Excel.Worksheet)
			End If
		End Function

		Friend Function pCellBorders(ByRef Cell As Excel.Range) As Excel.Border()
			Dim B As Excel.Borders
			B = Cell.Borders
			Dim Ret(5) As Excel.Border
			Ret(0) = B(Excel.XlBordersIndex.xlEdgeTop)
			Ret(1) = B(Excel.XlBordersIndex.xlEdgeRight)
			Ret(2) = B(Excel.XlBordersIndex.xlEdgeBottom)
			Ret(3) = B(Excel.XlBordersIndex.xlEdgeLeft)
			If Cell.EntireColumn.Count > 1 Then Ret(4) = B(Excel.XlBordersIndex.xlInsideVertical)
			If Cell.EntireRow.Count > 1 Then Ret(5) = B(Excel.XlBordersIndex.xlInsideHorizontal)
			Return Ret
		End Function

#Region "手順"
		''' <summary>Excelを起動します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 起動()
			GetExcelApp()
			ExcelApp.Visible = True
			Tips.SetForegroundWindow(CType(ExcelApp.Hwnd, IntPtr))
		End Sub
		Private Function GetExcelApp() As Excel.Application
			On Error Resume Next
			If StaticExcelApp Is Nothing Then
				Dim IsNotRunning As Boolean
				ExcelApp = TryCast(GetObject(, "Excel.Application"), Excel.Application)
				IsNotRunning = (Err().Number <> 0)
				Err.Clear()
				If IsNotRunning Then
					ExcelApp = New Excel.Application()
					Use(ExcelApp)
				End If
			Else
				ExcelApp = StaticExcelApp
			End If
			Init()
			Return ExcelApp
		End Function
		Private Sub Init()
			BookList = New WorkbookCollection(Me, ExcelApp.Workbooks)
			Use(BookList)
		End Sub

		''' <summary>エクセルの画面を表示します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 表示()
			EnsureApp()
			ExcelApp.Visible = True
		End Sub

		''' <summary>エクセルの画面を隠します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 非表示()
			非表示()
		End Sub

		''' <summary>エクセルの画面を隠します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 隠す()
			EnsureApp()
			ExcelApp.Visible = False
		End Sub

		''' <summary>すでに起動しているエクセルとリンクして、プロデルから利用できるようにします。</summary>
		<自分を>
		Public Sub リンク()
			If StaticExcelApp Is Nothing Then
				Try
					ExcelApp = TryCast(GetObject(, "Excel.Application"), Excel.Application)
					Use(ExcelApp)
				Catch When Err.Number = 429
					Throw New ProduireException("エクセルが起動していないためリンクできません。", ERRORBASE + 2)
				End Try
			Else
				ExcelApp = StaticExcelApp
			End If

			Init()
		End Sub

		''' <summary>Excelを起動して新しいブックを開きます</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Function 開く() As ブック
			起動()
			Return ワークブックを追加()
		End Function

		''' <summary>ワークブック(xls)を開きます</summary>
		'''　<remarks>【ファイル名】を</remarks>
		<自分で>
		Public Function 開く(<を> ByVal ファイル名 As String) As ブック
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください", ERRORBASE + 3)
			End If
			GetExcelApp()

			Dim Current As Excel.Workbook
			With ExcelApp
				Dim FullPath As String
				If Mid(ファイル名, 2, 1) <> ":" Then FullPath = FileUtils.YenSuffix(CurDir()) & ファイル名 Else FullPath = ファイル名
				Err.Clear()
				Try
					Current = BookList.Workbooks.Open(FullPath)
					Use(Current)
					ExcelApp.Visible = True
				Catch E As Exception
					Throw New ProduireException(E.Message, ERRORBASE + 3)
				End Try
			End With
			Dim ワークブック As ブック = New ブック(Me, Current)
			Return ワークブック
		End Function

		''' <summary>現在選択しているワークブックの内容を保存します。
		''' [ファイル名]を省略すると、上書き保存されます。</summary>
		'''　<remarks>【ファイル名】へ</remarks>
		<自分を>
		Public Sub 保存()
			選択ブック.保存()
		End Sub

		''' <summary>現在選択しているワークブックの内容を保存します。</summary>
		'''　<remarks>【ファイル名】へ</remarks>
		<自分を>
		Public Sub 保存(<へ> ByVal ファイル名 As String)
			選択ブック.保存(ファイル名)
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("PDFで"), 動詞("保存")>
		Public Sub PDFで保存(<へ> ByVal ファイル名 As String)
			選択ブック.PDFで保存(ファイル名)
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("XPSで"), 動詞("保存")>
		Public Sub XPSで保存(<へ> ByVal ファイル名 As String)
			選択ブック.XPSで保存(ファイル名)
		End Sub

		''' <summary>新しいワークブックを作成します。</summary>
		'''　<remarks>【対象】を</remarks>
		''' <example>
		''' エクセルに新しいワークブックを作成
		''' </example>
		<自分へ, 補語("新しいワークブックを"), 動詞("作成")>
		Public Function 新しいワークブックを作成() As ブック
			Return ワークブックを追加()
		End Function

		''' <summary>新しいワークブックを作成します。</summary>
		''' <example>
		''' エクセルのブックを作成
		''' </example>
		<自分へ, 補語("ブックを"), 動詞("追加", "作成")>
		Public Function ブックを追加() As ブック
			Return ワークブックを追加()
		End Function

		''' <summary>新しいワークブックを作成します。</summary>
		''' <example>
		''' エクセルのワークブックを作成
		''' </example>
		<自分へ, 補語("ワークブックを"), 動詞("追加", "作成")>
		Public Function ワークブックを追加() As ブック
			Dim NewBook As ブック
			NewBook = New ブック(Me, BookList.Workbooks.Add())
			BookList.Add(NewBook)
			NewBook.MyBook.Activate()
			Return NewBook
		End Function

		''' <summary>新しいワークシートを作成します。</summary>
		''' <example>
		''' エクセルに新しいワークシートを作成
		''' </example>
		<自分へ, 補語("新しいワークシートを"), 動詞("作成")>
		Public Function 新しいワークシートを作成() As シート
			Return 選択ブック.追加()
		End Function

		''' <summary>新しいワークシートを作成します。</summary>
		''' <example>
		''' エクセルのシートを作成
		''' </example>
		<自分へ, 補語("シートを"), 動詞("追加", "作成")>
		Public Function シートを追加() As シート
			Return ワークシートを追加()
		End Function

		''' <summary>新しいワークシートを作成します。</summary>
		''' <example>
		''' エクセルにワークシートを作成
		''' </example>
		<自分へ, 補語("ワークシートを"), 動詞("追加", "作成")>
		Public Function ワークシートを追加() As シート
			Return 選択ブック.追加()
		End Function

		''' <summary>保存確認せずにExcelを終了します</summary>
		<自分を>
		Public Sub 強制終了()
			ExcelApp.ActiveWorkbook.Saved = True
			終了()
		End Sub

		''' <summary>Excelを終了します</summary>
		<自分を>
		Public Sub 終了()
			ExcelApp.Quit()
			解放()
		End Sub

		<自分を>
		Private Sub 解放()
			BookList.Clear()
			BookList.Dispose()
			GC.Collect()
			Dispose()
			ExcelApp = Nothing
		End Sub

		''' <summary>ワークブックに含まれるマクロを実行します</summary>
		''' <remarks>【マクロ名】を〈【引数】で〉</remarks>
		<自分("にある")>
		Public Function マクロ実行(<を> ByVal マクロ名 As String, <で(), 省略()> ByVal 引数 As String()) As Object
			Dim Arr As New List(Of Object)
			Dim Result As Object
			Arr.Add(マクロ名)
			If Not 引数 Is Nothing Then Arr.AddRange(引数)
			Result = RunMacro(Arr.ToArray())
			If Err.Number <> 0 Then Throw New ProduireException(Err.Description, ERRORBASE + 7) : Err.Clear()
			Return Result
		End Function

		''' <summary>現在選択しているワークブックの内容を印刷します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 印刷()
			Dim MySheet As Excel.Worksheet = CType(ExcelApp.ActiveSheet, Excel.Worksheet)
			Use(MySheet)
			MySheet.PrintOut()
		End Sub

		''' <summary>現在選択しているワークブックの内容を印刷見本を表示します</summary>
		''' <remarks></remarks>
		<自分を>
		Public Sub プレビュー()
			Dim MySheet As Excel.Worksheet = CType(ExcelApp.ActiveSheet, Excel.Worksheet)
			Use(MySheet)
			MySheet.PrintPreview()
		End Sub

		''' <summary>現在選択しているワークシートの特定の範囲を選択状態にします</summary>
		'''　<remarks>【範囲】を</remarks>
		<自分へ>
		Public Sub 範囲選択(<を> ByVal Cell As String)
			ExcelApp.Range(Cell).Select()
			If Err.Number = 1004 Then Throw New ProduireException("セルの範囲指定が正しくありません。\n%1", ERRORBASE + 8) ', Cell
		End Sub

		''' <summary>エクセルが現在起動しているかどうかを表します</summary>
		<自分("が")>
		Public Function 起動中() As Boolean
			On Error Resume Next
			If ExcelApp Is Nothing Then ExcelApp = CType(GetObject(, "Excel.Application"), Excel.Application)
			Return (Not ExcelApp Is Nothing)
		End Function

#End Region

#Region "設定項目"
		''' <summary>エクセルのバージョン情報を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property バージョン() As String
			Get
				With ExcelApp
					Return .Name & " " & .Version
				End With
			End Get
		End Property

		''' <summary>Office連携プラグインのバージョン情報を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property プラグインのバージョン() As String
			Get
				Return My.Application.Info.Title & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
			End Get
		End Property

		''' <summary>現在選択しているワークブックのファイル名を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property ファイル名() As String
			Get
				Return ExcelApp.ActiveWorkbook.FullName
			End Get
		End Property

		''' <summary>現在開いているワークブックの一覧を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property ワークブック一覧() As String()
			Get
				If Not ExcelApp Is Nothing Then
					Dim Arr As String()
					With GetExcelApp()
						Dim I As Integer
						Dim TmpWorkbook As Excel.Workbook
						Dim TmpWorkbooks As Excel.Workbooks = .Workbooks
						Use(TmpWorkbooks)
						ReDim Arr(TmpWorkbooks.Count - 1)
						For I = 1 To TmpWorkbooks.Count
							TmpWorkbook = TmpWorkbooks.Item(I)
							Arr(I - 1) = TmpWorkbook.Name
							Use(TmpWorkbook)
						Next
					End With
					Return Arr
				Else
					Return Nothing
				End If
			End Get
		End Property

		''' <summary>現在選択しているワークブックを表します</summary>
		''' <returns>□</returns>
		Public Property ワークブック() As String
			Get
				Dim Book As Excel.Workbook
				Book = ExcelApp.ActiveWorkbook
				Return Book.Name
			End Get
			Set(ByVal Value As String)
				If Value.Length = 0 Then Exit Property
				Dim Books As Excel.Workbooks, Book As Excel.Workbook
				Books = GetExcelApp().Workbooks
				Try
					Book = Books(CStr(Value))
				Catch ex As Exception
					Throw New ProduireException(ex)
				End Try
				Book.Activate()
			End Set
		End Property

		''' <summary>現在選択しているワークシートを表します</summary>
		''' <returns>□</returns>
		Public Property ワークシート() As String
			Get
				Dim Book As Excel.Workbook, Sheet As Excel.Worksheet
				Book = ExcelApp.ActiveWorkbook
				Sheet = CType(Book.ActiveSheet, Excel.Worksheet)
				Return Sheet.Name
			End Get
			Set(ByVal value As String)
				Dim Book As Excel.Workbook, Sheets As Excel.Sheets, Sheet As Excel.Worksheet
				Book = ExcelApp.ActiveWorkbook
				Sheets = Book.Sheets
				Sheet = CType(Sheets.Item(value), Excel.Worksheet)
				Sheet.Activate()
			End Set
		End Property

		''' <summary>現在開いているワークブックの数を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property ワークブック数() As Integer
			Get
				Return BookList.Count
			End Get
		End Property

		''' <summary>エクセルが画面に表示されているかどうかを表します</summary>
		Public Property 表示状態() As Boolean
			Get
				EnsureApp()
				Return ExcelApp.Visible
			End Get
			Set(ByVal value As Boolean)
				EnsureApp()
				ExcelApp.Visible = value
			End Set
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		Public ReadOnly Property 選択セル() As セル
			Get
				Dim ActiveCell As Excel.Range = ExcelApp.ActiveCell
				If ActiveCell Is Nothing Then Return Nothing

				Dim Cell As New セル(Me, ExcelApp.ActiveCell)
				Return Cell
			End Get
		End Property

		''' <summary></summary>
		''' <returns>◎</returns>
		Public Property 警告表示() As Boolean
			Get
				If Not EnsureApp() Then Return True
				Return ExcelApp.DisplayAlerts
			End Get
			Set(ByVal value As Boolean)
				If ExcelApp Is Nothing Then Throw New ProduireException("エクセルを起動またはリンクしていないためリンクできません。", ERRORBASE + 2) : Return
				ExcelApp.DisplayAlerts = value
			End Set
		End Property

		Private Function EnsureApp() As Boolean
			If ExcelApp Is Nothing Then Throw New ProduireException("エクセルを起動またはリンクしていないためリンクできません。", ERRORBASE + 2)
		End Function

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("選択ブック", "現在ブック")>
		Public ReadOnly Property 選択ブック() As ブック
			Get
				Dim ActiveWorkbook As Excel.Workbook = GetExcelApp().ActiveWorkbook
				If ActiveWorkbook Is Nothing Then Return Nothing

				Dim Book As New ブック(Me, ActiveWorkbook)
				Return Book
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("選択シート", "現在シート")>
		Public ReadOnly Property 選択シート() As シート
			Get
				Dim ActiveSheet As Excel.Worksheet = CType(GetExcelApp().ActiveSheet, Excel.Worksheet)
				If ActiveSheet Is Nothing Then Return Nothing

				Dim Sheet As New シート(Me, ActiveSheet)
				Use(Sheet)
				Return Sheet
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Excel.Application
			Get
				Return ExcelApp
			End Get
		End Property

#End Region

		Private Function RunMacro(ByVal oRunArgs As Object()) As Object
			Return GetType(Excel.Application).InvokeMember("Run", System.Reflection.BindingFlags.Default Or System.Reflection.BindingFlags.InvokeMethod, Nothing, ExcelApp, oRunArgs)
		End Function

#Region "IObjectContainer"
		Public ReadOnly Property 子部品一覧() As Object() Implements IObjectContainer.子部品一覧
			Get
				If BookList Is Nothing Then Return Nothing
				Dim Items(BookList.Count) As Object
				BookList.CopyTo(CType(Items, ブック()))
				Return Items
			End Get
		End Property
		<除外>
		Public Overloads Function TryGeyValue(ByVal Key As String, ByRef Value As IProduireClass) As Boolean Implements IObjectContainer.TryGeyValue
			If BookList Is Nothing Then Return False
			Dim Result As Boolean = BookList.Contains(Key)
			If Result Then Value = BookList(Key)
			Return Result
		End Function

		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub

		Public ReadOnly Property Name() As String Implements IObjectContainer.Name
			Get
				Return ""
			End Get
		End Property

#End Region

	End Class

	Public Class WorkbookCollection
		Inherits List(Of ブック)
		Implements IDisposable

		ReadOnly ExcelApp As エクセル
		Public Workbooks As Excel.Workbooks

		Public Sub New(ByVal ExcelApp As エクセル, ByVal Books As Excel.Workbooks)
			Me.ExcelApp = ExcelApp
			Me.Workbooks = Books
		End Sub

		Public Overloads Function Contains(ByVal key As String) As Boolean
			Dim Book As Excel.Workbook
			For Each Book In Workbooks
				If Book.FullName = key Then
					Return True
					Exit For
				End If
			Next
			Return False
		End Function

		Public Overloads Sub Clear()
			Dim MyBook As ブック
			For Each MyBook In Me
				MyBook.Dispose()
			Next
			MyBase.Clear()
		End Sub

		Public Sub Refresh()
			Dim Book As Excel.Workbook
			For Each Book In Workbooks
				Me.Add(New ブック(ExcelApp, Book))
			Next
		End Sub

		Protected Overridable Sub Dispose(ByVal disposing As Boolean)
			If Not Workbooks Is Nothing Then
				If disposing Then
					Me.Clear()
					Marshal.FinalReleaseComObject(Workbooks)
					Workbooks = Nothing
				End If
			End If
		End Sub

#Region " IDisposable Support "
		' このコードは、破棄可能なパターンを正しく実装できるように Visual Basic によって追加されました。
		Public Sub Dispose() Implements IDisposable.Dispose
			' このコードを変更しないでください。クリーンアップ コードを上の Dispose(ByVal disposing As Boolean) に記述します。
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub
#End Region

		Protected Overrides Sub Finalize()
			Dispose()
			MyBase.Finalize()
		End Sub

		Default Public Overloads ReadOnly Property Item(ByVal key As String) As ブック
			Get
				On Error Resume Next
				Dim Arr() As String
				Dim I As Integer, Book As Excel.Workbook

				If IsNumeric(key) Then
					Book = Workbooks(CShort(key))
				Else
					Book = Workbooks(key)
				End If
				Return New ブック(ExcelApp, Book)
			End Get
		End Property
	End Class

	''' <summary>
	''' Microsoft Excelのワークブックに関する情報です。
	''' ワークブックとは、Excelのファイル一つ一つのことを表しています。
	''' つまり、Excelのファイルを開くと１つのワークブックが表示されます。
	''' ワークブックには、複数のワークシートが含まれています。
	''' </summary>
	<種類(DocUrl:="/office/workbook.htm")>
	Public Class ブック
		Inherits DisposableObject
		Implements IProduireClass

		Dim ExcelApp As Excel.Application
		Public MyBook As Excel.Workbook
		ReadOnly RdrExcel As エクセル

		Public Sub New(ByVal RdrExcel As エクセル, ByVal MyBook As Excel.Workbook)
			ExcelApp = MyBook.Application
			Me.MyBook = MyBook
			Me.RdrExcel = RdrExcel
			Use(MyBook)
		End Sub

#Region "手順"
		''' <summary>ワークシートを選択します</summary>
		'''　<remarks>【シート名】を</remarks>
		<自分で>
		Public Sub 選択(<を> ByVal SheetName As String)
			Dim Sheet As Excel.Worksheet
			Sheet = CType(MyBook.Sheets.Item(SheetName), Excel.Worksheet)
			Sheet.Activate()
		End Sub

		''' <summary>現在のワークブックにシートを追加します</summary>
		'''　<remarks></remarks>
		<自分へ, 補語("シートを"), 動詞("追加", "作成")>
		Public Function 追加() As シート
			Dim NewSheet As Excel.Worksheet
			NewSheet = CType(ExcelApp.Worksheets.Add(), Excel.Worksheet)
			Use(NewSheet)
			NewSheet.Activate()

			Dim PSheet As New シート(RdrExcel, NewSheet)
			Return PSheet
		End Function

		''' <summary>ワークブックを指定したファイル名で保存します</summary>
		<自分を>
		Public Sub 保存()
			'On Error Resume Next
			With MyBook
				.Save()
				'.Saved = True
				'.Application.Visible = True
			End With
		End Sub

		''' <summary>ワークブックを上書き保存します</summary>
		<自分を>
		Public Sub 上書き保存()
			'On Error Resume Next
			Dim Alerts As Boolean = ExcelApp.DisplayAlerts
			ExcelApp.DisplayAlerts = False
			MyBook.SaveAs()
			ExcelApp.DisplayAlerts = Alerts
		End Sub

		''' <summary>ワークブックを指定したファイル名で保存します</summary>
		'''　<remarks>【ファイル名】へ</remarks>
		<自分を>
		Public Sub 保存(<へ> ByVal ファイル名 As String)
			'On Error Resume Next
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyBook
				.SaveAs(ファイル名)
				'.Saved = True
				'.Application.Visible = True
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("PDFで"), 動詞("保存")>
		Public Sub PDFで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyBook
				.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, ファイル名)
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("XPSで"), 動詞("保存")>
		Public Sub XPSで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyBook
				.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypeXPS, ファイル名)
			End With
		End Sub

		''' <summary>ワークブックを閉じます</summary>
		<自分を>
		Public Sub 閉じる()
			MyBook.Close()
		End Sub

		''' <summary>ワークブックを閉じます</summary>
		<自分を, 補語("保存せずに")>
		Public Sub 保存せずに閉じる()
			MyBook.Close(False)
		End Sub

		''' <summary>ワークブックを印刷します</summary>
		<自分を>
		Public Sub 印刷()
			MyBook.PrintOut()
		End Sub

		''' <summary>ワークブックのプレビューを表示します</summary>
		<自分を>
		Public Sub プレビュー()
			MyBook.Sheets.PrintPreview()
		End Sub

		''' <summary>ワークブックに含まれるマクロを実行します
		''' 他のブックのマクロを実行する場合、下記参照</summary>
		'''　<remarks>【マクロ名】を〈【引数】で〉</remarks>
		<自分("にある")>
		Public Function マクロ実行(<を> ByVal MacroName As String, <で(), 省略()> ByVal Paramater As String()) As Object
			Dim Arr As New List(Of Object)
			Dim Result As Object
			Arr.Add(MacroName)
			If Not Paramater Is Nothing Then Arr.AddRange(Paramater)
			Result = RunMacro(Arr.ToArray())
			If Err.Number <> 0 Then Throw New ProduireException(Err.Description, エクセル.ERRORBASE + 7) : Err.Clear()
			Return Result
		End Function

		Private Function RunMacro(ByVal oRunArgs As Object()) As Object
			Return GetType(Excel.Application).InvokeMember("Run", System.Reflection.BindingFlags.Default Or System.Reflection.BindingFlags.InvokeMethod, Nothing, ExcelApp, oRunArgs)
		End Function
#End Region

#Region "設定項目"
		''' <summary>ワークブックのファイル名</summary>
		'''// <remarks>□</remarks>
		Public ReadOnly Property ファイル名() As String
			Get
				Return MyBook.FullName
			End Get
		End Property

		Public Property ワークシート() As String
			Get
				Dim Sheet As Excel.Worksheet
				Sheet = CType(MyBook.ActiveSheet, Excel.Worksheet)
				Return Sheet.Name
			End Get
			Set(ByVal value As String)
				Dim Sheets As Excel.Sheets, Sheet As Excel.Worksheet
				Sheets = MyBook.Sheets
				Try
					Sheet = CType(Sheets.Item(value), Excel.Worksheet)
				Catch ex As Exception
					Throw New ProduireException(ex)
				End Try
				Sheet.Activate()
			End Set
		End Property

		''' <summary>ワークブックに含まれているワークシートの一覧</summary>
		'''// <remarks>□</remarks>
		Public ReadOnly Property ワークシート一覧() As String()
			Get
				Dim Result As String()
				With MyBook
					Dim TmpSheet As Excel.Worksheet
					ReDim Result(.Sheets.Count - 1)
					Dim I As Integer
					For I = 1 To .Sheets.Count
						TmpSheet = CType(.Sheets(I), Excel.Worksheet)
						Result(I - 1) = TmpSheet.Name
					Next
				End With
				Return Result
			End Get
		End Property

		''' <summary>ワークブックに含まれているワークシートの数</summary>
		'''// <remarks>□</remarks>
		Public ReadOnly Property ワークシート数() As Integer
			Get
				Return MyBook.Sheets.Count
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("選択シート", "現在シート")>
		Public ReadOnly Property 選択シート() As シート
			Get
				Dim ActiveSheet As Excel.Worksheet = CType(MyBook.ActiveSheet, Excel.Worksheet)
				If ActiveSheet Is Nothing Then Return Nothing

				Dim PSheet As New シート(RdrExcel, ActiveSheet)
				Use(PSheet)
				Return PSheet
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("末尾シート")>
		Public ReadOnly Property 末尾シート() As シート
			Get
				Dim LastSheet As Excel.Worksheet = CType(MyBook.Worksheets(MyBook.Worksheets.Count), Excel.Worksheet)
				If LastSheet Is Nothing Then Return Nothing
				Dim PSheet As New シート(RdrExcel, LastSheet)
				Use(PSheet)
				Return PSheet
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("先頭シート")>
		Public ReadOnly Property 先頭シート() As シート
			Get
				Dim FirstSheet As Excel.Worksheet = CType(MyBook.Worksheets(1), Excel.Worksheet)
				If FirstSheet Is Nothing Then Return Nothing
				Dim PSheet As New シート(RdrExcel, FirstSheet)
				Use(PSheet)
				Return PSheet
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Excel.Workbook
			Get
				Return MyBook
			End Get
		End Property

#End Region

	End Class

	''' <summary>
	''' Microsoft Excelのワークシートに関する情報です。
	''' ワークシートは、Excelのファイルに複数含まれています。
	''' 通常は、ウィンドウの下のタブで切り替えることができます。
	''' </summary>
	<種類(DocUrl:="/office/worksheet.htm")>
	Public Class シート
		Inherits DisposableObject
		Implements IProduireClass

		Dim ExcelApp As Excel.Application
		Dim MySheet As Excel.Worksheet
		Dim RdrExcel As エクセル

		Public Sub New(ByVal RdrExcel As エクセル, ByVal MySheet As Excel.Worksheet)
			Me.RdrExcel = RdrExcel
			Me.MySheet = MySheet
			Me.ExcelApp = MySheet.Application
			Use(MySheet)
		End Sub

#Region "手順"

		''' <summary>指定した範囲のセルを選択します</summary>
		<自分を>
		Public Sub 選択する()
			MySheet.Activate()
		End Sub

		''' <summary>指定した範囲のセルを選択します</summary>
		'''　<remarks>【開始セル】から【終了セル】まで</remarks>
		<自分で>
		Public Function 選択する(<から> ByVal FromCell As String, <まで> ByVal ToCell As String) As セル
			On Error Resume Next
			Dim Cell As Excel.Range
			Cell = MySheet.Cells.Range(FromCell, ToCell)
			If Cell Is Nothing Then
				Throw New ProduireException("指定したセルは見つかりません。", エクセル.ERRORBASE + 2)
			End If
			Dim MyCell As New セル(RdrExcel, Cell)
			MyCell.選択する()
			Return MyCell
		End Function

		''' <summary>指定した列を選択します</summary>
		<自分で>
		Public Function 列選択する(<を> ByVal Name As String) As セル
			On Error Resume Next
			Dim Cell As Excel.Range
			Cell = CType(MySheet.Columns(Name), Excel.Range)
			If Cell Is Nothing Then
				Throw New ProduireException("指定したセルは見つかりません。", エクセル.ERRORBASE + 2)
			End If
			Dim MyCell As New セル(RdrExcel, Cell)
			MyCell.選択する()
			Return MyCell
		End Function

		''' <summary>指定した行を選択します</summary>
		<自分で>
		Public Function 行選択する(<を> ByVal Name As String) As セル
			On Error Resume Next
			Dim Cell As Excel.Range
			Cell = CType(MySheet.Rows(Name), Excel.Range)
			If Cell Is Nothing Then
				Throw New ProduireException("指定したセルは見つかりません。", エクセル.ERRORBASE + 2)
			End If
			Dim MyCell As New セル(RdrExcel, Cell)
			MyCell.選択する()
			Return MyCell
		End Function

		''' <summary>セルをすべて選択します</summary>
		<自分を>
		Public Sub すべて選択()
			Dim Range As New セル(RdrExcel, MySheet.UsedRange)
			Range.選択する()
		End Sub

		''' <summary>シートを削除します</summary>
		<自分を>
		Public Sub 削除()
			MySheet.Delete()
		End Sub

		''' <summary>シートを《コピー先》のシートの後ろへコピーします</summary>
		'''　<remarks>《コピー先》へ</remarks>
		<自分を>
		Public Function コピー(<へ()> ByVal Value As String) As シート
			MySheet.Copy(, RdrExcel.pGetWorkSheet(Value))
			Return RdrExcel.選択シート
		End Function

		''' <summary>シートを《コピー先》のシートの後ろへコピーします</summary>
		'''　<remarks>《コピー先》へ</remarks>
		<自分を>
		Public Function コピー() As シート
			MySheet.Copy(, MySheet)
			Return RdrExcel.選択シート
		End Function

		''' <summary>シートを《移動先》のシートの後ろに移動します</summary>
		'''　<remarks>《移動先》へ</remarks>
		<自分を>
		Public Sub 移動(<へ> ByVal Value As String)
			If Len(Value) <> 0 Then
				MySheet.Move(, RdrExcel.pGetWorkSheet(Value))
			Else
				MySheet.Move()
			End If
		End Sub

		''' <summary>シートを《移動先》のシートの後ろに移動します</summary>
		'''　<remarks>《移動先》へ</remarks>
		<自分を>
		Public Sub 移動()
			MySheet.Move()
		End Sub

		''' <summary>シートを《移動先》のシートの後ろに移動します</summary>
		'''　<remarks>《移動先》へ</remarks>
		<自分を, 補語("後へ"), 動詞("移動")>
		Public Sub 後へ移動(<助詞("より")> ByVal Sheet As シート)
			MySheet.Move(, Sheet.MySheet)
		End Sub

		''' <summary>シートを《移動先》のシートの後ろに移動します</summary>
		'''　<remarks>《移動先》へ</remarks>
		<自分を, 補語("前へ"), 動詞("移動")>
		Public Sub 前へ移動(<助詞("より")> ByVal Sheet As シート)
			MySheet.Move(Sheet.MySheet)
		End Sub

		''' <summary>シートを保護します</summary>
		<自分を>
		Public Sub 保護する()
			MySheet.Protect()
		End Sub

		''' <summary>シートをパスワードで保護します</summary>
		<自分を>
		Public Sub 保護する(<で> ByVal パスワード As String)
			MySheet.Protect(Password:=パスワード)
		End Sub

		''' <summary>シートの保護を解除します</summary>
		<自分を>
		Public Sub 解除する()
			MySheet.Unprotect()
		End Sub

		''' <summary>シートの保護を解除します</summary>
		<自分を>
		Public Sub 解除する(<で> ByVal パスワード As String)
			MySheet.Unprotect(Password:=パスワード)
		End Sub

		''' <summary>シートを再計算します</summary>
		<自分を>
		Public Sub 計算()
			MySheet.Calculate()
		End Sub

		''' <summary>シートを印刷</summary>
		<自分を>
		Public Sub 印刷()
			MySheet.PrintOut()
		End Sub

		''' <summary>シートの印刷見本を表示します</summary>
		<自分を>
		Public Sub プレビュー()
			MySheet.PrintPreview()
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("PDFで"), 動詞("保存")>
		Public Sub PDFで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MySheet
				.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, ファイル名)
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("XPSで"), 動詞("保存")>
		Public Sub XPSで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MySheet
				.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypeXPS, ファイル名)
			End With
		End Sub

		''' <summary>オートフィルタで、【列番号】を【条件式】に一致するものだけを表示させます</summary>
		'''　<remarks>【列番号】を【条件式】で</remarks>
		Public Sub フィルタ(<を> ByVal 列番号 As Integer, <で()> ByVal 条件1 As String)
			RdrExcel.選択セル.MyRange.AutoFilter(列番号, 条件1)
		End Sub

		''' <summary>セルを選択します</summary>
		'''　<remarks>｛《横》,《縦》｝の
		''' 【セル名】の</remarks>
		<名詞手順>
		Public Function セル(<既定()> ByVal Pos As String) As セル
			On Error Resume Next
			Dim Cell As Excel.Range
			Cell = MySheet.Cells.Range(Pos)
			If Cell Is Nothing Then
				Throw New ProduireException("指定したセル'" + Pos(0) + "'は見つかりません。", エクセル.ERRORBASE + 2)
				Return Nothing
			End If

			Dim MyCell As New セル(RdrExcel, Cell)
			Use(MyCell)
			Return MyCell
		End Function

		''' <summary>セルを選択します</summary>
		'''　<remarks>｛《横》,《縦》｝の
		''' 【セル名】の</remarks>
		<名詞手順>
		Public Function セル(<既定()> ByVal RowCol As Object()) As セル
			On Error Resume Next
			Dim Cell As Excel.Range
			If RowCol.Length = 2 Then
				Cell = CType(MySheet.Cells.Item(Integer.Parse(RowCol(0).ToString()), Integer.Parse(RowCol(1).ToString())), Excel.Range)
				If Cell Is Nothing Then
					Throw New ProduireException("指定したセル'" + RowCol(0).ToString() + "," + RowCol(1).ToString() + "'は見つかりません。", エクセル.ERRORBASE + 2)
					Return Nothing
				End If
			ElseIf RowCol.Length = 4 Then
				Cell = CType(MySheet.Cells.Item(Integer.Parse(RowCol(0).ToString()), Integer.Parse(RowCol(1).ToString())), Excel.Range)
				If Cell Is Nothing Then
					Throw New ProduireException("指定したセル'" + RowCol(0).ToString() + "," + RowCol(1).ToString() + "'は見つかりません。", エクセル.ERRORBASE + 2)
					Return Nothing
				End If
				Dim Cell2 As Excel.Range = CType(MySheet.Cells.Item(Integer.Parse(RowCol(2).ToString()), Integer.Parse(RowCol(3).ToString())), Excel.Range)
				If Cell2 Is Nothing Then
					Throw New ProduireException("指定したセル'" + RowCol(2).ToString() + "," + RowCol(3).ToString() + "'は見つかりません。", エクセル.ERRORBASE + 2)
					Return Nothing
				End If
				Cell = MySheet.Cells.Range(Cell, Cell2)
			Else
				Cell = MySheet.Cells.Range(RowCol(0))
				If Cell Is Nothing Then
					Throw New ProduireException("指定したセル'" + RowCol(0).ToString() + "'は見つかりません。", エクセル.ERRORBASE + 2)
					Return Nothing
				End If
			End If

			Dim MyCell As New セル(RdrExcel, Cell)
			Use(MyCell)
			Return MyCell
		End Function
#End Region

#Region "図形関連"

		''' <summary>シートへ画像を追加します</summary>
		<自分へ, 手順名("画像として", "追加")>
		Public Function 画像として追加する(<を> ByVal Value As String) As エクセル図形
			Dim Shapes As Excel.Shapes = GetShapes()
			Dim MyShape As Excel.Shape = Shapes.AddPicture(Value, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, -1, -1)
			Use(MyShape)
			Return New エクセル図形(RdrExcel, MyShape)
		End Function

		''' <summary>指定した名前の図形を取得します</summary>
		''' <remarks>【名前】の</remarks>
		<名詞手順>
		Public Function 図形(<既定()> ByVal Pos As String) As エクセル図形
			On Error Resume Next
			Dim Shapes As Excel.Shapes = GetShapes()
			Dim MyShape As Excel.Shape = CType(Shapes.Range(CType(Pos, Object)), Excel.Shape)
			Use(MyShape)

			If MyShape Is Nothing Then
				Throw New ProduireException("指定した図形'" + Pos + "'は見つかりません。", エクセル.ERRORBASE + 2)
				Return Nothing
			End If
			Return New エクセル図形(RdrExcel, MyShape)
		End Function

		Private Function GetShapes() As Excel.Shapes
			Dim Shapes As Excel.Shapes = MySheet.Shapes
			Use(Shapes)
			Return Shapes
		End Function

#End Region

#Region "設定項目"
		''' <summary>ワークシートの名前</summary>
		Public Property 名前() As String
			Get
				Return MySheet.Name
			End Get
			Set(ByVal Value As String)
				MySheet.Name = Value
			End Set
		End Property

		''' <summary>ワークシート上にある図形の名前一覧を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 図形一覧() As String()
			Get
				Dim I As Integer
				Dim Shapes As Excel.Shapes = GetShapes()
				Dim TmpShape As Excel.Shape

				Dim Result As String()
				ReDim Result(Shapes.Count - 1)
				For I = 1 To Shapes.Count
					TmpShape = Shapes.Item(I)
					Result(I - 1) = TmpShape.Name
					Use(TmpShape)
				Next
				Return Result
			End Get
		End Property

		''' <summary>ワークシート上にある図形のコメント一覧を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property コメント一覧() As エクセルコメント()
			Get
				Dim I As Integer
				Dim Comments As Excel.Comments = MySheet.Comments
				Use(Comments)
				Dim TmpComment As Excel.Comment

				Dim Result As エクセルコメント()
				ReDim Result(Comments.Count - 1)
				For I = 1 To Comments.Count
					TmpComment = Comments.Item(I)
					Result(I - 1) = New エクセルコメント(RdrExcel, TmpComment)
					Use(TmpComment)
				Next
				Return Result
			End Get
		End Property

		''' <summary>ワークシートのすべてのセルの値
		''' 2次配列で表されます</summary>
		Public Property 一覧() As String()()
			Get
				Dim Arr()() As String
				Dim I As Integer, J As Integer
				Dim UsedRange As Excel.Range
				UsedRange = MySheet.UsedRange
				With UsedRange
					Dim ColumnCount As Integer = .Columns.Count
					Dim RowCount As Integer = .Rows.Count
					ReDim Arr(RowCount - 1)
					For I = 1 To RowCount
						ReDim Arr(I - 1)(ColumnCount - 1)
						For J = 1 To ColumnCount
							Dim Cell As Object = ._Default(I, J)
							If TypeOf Cell Is String Then
								Arr(I - 1)(J - 1) = CType(Cell, String)
							Else
								Dim Cell2 As Excel.Range = CType(Cell, Excel.Range)
								Arr(I - 1)(J - 1) = CType(Cell2.Text, String)
							End If
						Next J
					Next
				End With
				Return Arr
			End Get
			Set(ByVal value As String()())
				Dim Arr2 As String()
				Dim I As Integer, J As Integer, Count As Integer
				For I = 0 To value.GetUpperBound(0)
					Arr2 = value(I)
					Count = UBound(Arr2)
					For J = 0 To Count
						MySheet.Cells._Default(I + 1, J + 1) = Arr2(J)
					Next
				Next
			End Set
		End Property

		''' <summary>ワークシートのすべてのセルの値
		''' 2次配列で表されます</summary>
		Public Property 値一覧() As String()()
			Get
				Dim UsedRange As Excel.Range
				UsedRange = MySheet.UsedRange
				Dim Arr()() As String
				Dim I As Integer, J As Integer
				With UsedRange
					Dim ColumnCount As Integer = .Columns.Count
					Dim RowCount As Integer = .Rows.Count
					ReDim Arr(RowCount - 1)
					For I = 1 To RowCount
						ReDim Arr(I - 1)(ColumnCount - 1)
						For J = 1 To ColumnCount
							Dim Cell As Object = ._Default(I, J)
							If TypeOf Cell Is String Then
								Arr(I - 1)(J - 1) = CType(Cell, String)
							Else
								Dim Cell2 As Excel.Range = CType(Cell, Excel.Range)
								Arr(I - 1)(J - 1) = CType(Cell2.Value, String)
							End If
						Next J
					Next
				End With
				Return Arr
			End Get
			Set(ByVal value As String()())
				一覧 = value
			End Set
		End Property

		''' <summary>選択しているセルを表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 選択() As セル
			Get
				If TypeOf ExcelApp.Selection Is Excel.Range Then
					Return New セル(RdrExcel, CType(ExcelApp.Selection, Excel.Range))
				Else
					Return Nothing
				End If
			End Get
		End Property


		''' <summary>ワークシートの選択範囲
		''' A1:F1のようなR1C1形式で表します</summary>
		<設定項目("選択範囲", "範囲選択")> Public Property 選択範囲() As String
			Get
				Return CType(ExcelApp.Selection, String)
			End Get
			Set(ByVal value As String)
				MySheet.Range(value).Select()
				If Err.Number = 1004 Then Throw New ProduireException("セルの範囲指定が正しくありません。\n" + value, エクセル.ERRORBASE + 8)
			End Set
		End Property

		''' <summary>ワークシートで使用されているセル全体を表します</summary>
		''' <example>エクセルの現在シートの使用範囲の行数を表示</example>
		''' <returns>□</returns>
		Public ReadOnly Property 使用範囲() As セル
			Get
				Dim Cell As Excel.Range
				Cell = MySheet.UsedRange
				Return New セル(RdrExcel, Cell)
			End Get
		End Property

		''' <summary>オートフィルタを有効にするかどうか</summary>
		Public Property オートフィルタ() As Boolean
			Get
				Return MySheet.AutoFilterMode
			End Get
			Set(ByVal value As Boolean)
				MySheet.AutoFilterMode = value
			End Set
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Excel.Worksheet
			Get
				Return MySheet
			End Get
		End Property

#End Region

		Friend Function GetShape(ByRef ShapeName As String) As Excel.Shape
			Dim Shapes As Excel.Shapes = GetShapes()
			Return Shapes.Item(CStr(ShapeName))
		End Function

	End Class
End Namespace
