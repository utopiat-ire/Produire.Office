' プロデル Microsoft Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Strict On
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports utopiat
Imports Microsoft.Office.Core
Imports System.Reflection

Namespace ワード

	''' <summary>ワード</summary>
	<種類(DocUrl:="/office/word.htm")>
	Public Class ワード
		Inherits DisposableObject
		Implements IProduireStaticClass
		Implements IObjectContainer

		Private WordApp As Word.Application
		Public Const ERRORBASE As Short = 2100

#Region "手順"
		'ワード起動	{=1}Aで	可視A(オンかオフ)でワードを起動する
		''' <summary>Wordを起動します</summary>
		''' <remarks></remarks>
		<自分を>
		Public Sub 起動()
			GetWordApp()
			WordApp.Visible = True
		End Sub

		Private Sub GetWordApp()
			On Error Resume Next
			Dim IsNotRunning As Boolean
			WordApp = DirectCast(GetObject(, "Word.Application"), Word.Application)
			IsNotRunning = (WordApp Is Nothing)
			Err.Clear()
			If IsNotRunning Then
				WordApp = New Word.Application()
				Use(WordApp)
			End If

			Init()
		End Sub

		''' <summary>エクセルの画面を隠します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Sub 隠す()
			If WordApp Is Nothing Then Throw New ProduireException("エクセルを起動またはリンクしていないためリンクできません。", ERRORBASE + 2)
			WordApp.Visible = False
		End Sub

		''' <summary>すでに起動しているワードとリンクして、プロデルから利用できるようにします。</summary>
		<自分("へ", "を")>
		Public Sub リンク()
			If Not WordApp Is Nothing Then Return
			Try
				WordApp = DirectCast(GetObject(, "Word.Application"), Word.Application)
				Use(WordApp)
				Init()
			Catch When Err.Number = 429
				Throw New ProduireException("ワードが起動していないためリンクできません。", ERRORBASE + 2)
			End Try
		End Sub

		Private Sub Init()
			DocList = New DocumentCollection(Me, WordApp.Documents)
			Use(DocList)
		End Sub

		''' <summary>現在操作している文章を保存します</summary>
		<自分を>
		Public Sub 保存()
			Me.現在文章.保存()
		End Sub

		''' <summary>現在操作している文章を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を>
		Public Sub 保存(<へ> ByVal ファイル名 As String)
			Me.現在文章.保存(ファイル名)
		End Sub

		''' <summary>新しい文章を追加します</summary>
		''' <remarks>ドキュメントを</remarks>
		<自分へ, 補語("ドキュメントを"), 動詞("追加", "作成")>
		Public Function ドキュメントを追加() As 文章
			Return 新規文書を作成()
		End Function

		''' <summary>新しい文章を追加します</summary>
		''' <remarks>文章を</remarks>
		<自分へ, 補語("文章を"), 動詞("追加", "作成")>
		Public Function 文章を追加() As 文章
			Return 新規文書を作成()
		End Function

		''' <summary>新しい文章を追加します</summary>
		<自分へ, 補語("新規文書を"), 動詞("追加", "作成")>
		Public Function 新規文書を作成() As 文章
			Dim NewDoc As 文章
			NewDoc = New 文章(Me, WordApp.Documents.Add())
			NewDoc.MyDoc.Activate()
			Return NewDoc
		End Function


		''' <summary>Wordを起動して新しいブックを開きます</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Function 開く() As 文章
			起動()
			Return 新規文書を作成()
		End Function

		''' <summary>指定したファイルを開きます</summary>
		''' <remarks>【ファイル名】を</remarks>
		<自分で>
		Public Function 開く(<を> ByVal ファイル名 As String) As 文章
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください", ERRORBASE + 3)
			End If
			GetWordApp()

			Dim FileName2 As String
			Dim Current As Word.Document
			With WordApp
				If Mid(ファイル名, 2, 1) <> ":" Then FileName2 = FileUtils.YenSuffix(CurDir()) & ファイル名 Else FileName2 = ファイル名
				Err.Clear()
				Try
					Current = .Documents.Open(CType(ファイル名, Object))
					Use(Current)
					WordApp.Visible = True
				Catch E As Exception
					Throw New ProduireException(E.Message, ERRORBASE + 3)
				End Try
			End With
			Dim 文章 As 文章 = New 文章(Me, Current)
			Return 文章
		End Function

		''' <summary>保存確認せずにWordを終了します</summary>
		<自分を>
		Public Sub 強制終了()
			WordApp.ActiveDocument.Saved = True
			終了()
		End Sub

		''' <summary>Wordを終了します</summary>
		''' <remarks></remarks>
		<自分を>
		Public Sub 終了()
			WordApp.Quit()
			解放()
		End Sub

		Private Sub 解放()
			DocList.Clear()
			DocList.Dispose()
			GC.Collect()
			Dispose()
			WordApp = Nothing
		End Sub

		''' <summary>文章に含まれるマクロを実行します</summary>
		'''　<remarks>【関数名】を〈【引数】で〉</remarks>
		<自分("にある")>
		Public Function マクロ実行(<を> ByVal マクロ名 As String, <で()> ByVal 引数 As String()) As Object
			Dim Arr As New List(Of Object)
			Dim Result As Object
			Arr.Add(マクロ名)
			If Not 引数 Is Nothing Then Arr.AddRange(引数)
			Result = RunMacro(Arr.ToArray())
			If Err.Number <> 0 Then Throw New ProduireException(Err.Description, ERRORBASE + 7) : Err.Clear()
			Return Result
		End Function

		Private Function RunMacro(ByVal oRunArgs As Object()) As Object
			Return GetType(Word.Application).InvokeMember("Run", Reflection.BindingFlags.Default Or Reflection.BindingFlags.InvokeMethod, Nothing, WordApp, oRunArgs)
		End Function
#End Region

#Region "設定項目"

		''' <summary>Microsoft Wordのバージョン</summary>
		''' <returns>○</returns>
		Public ReadOnly Property バージョン() As String
			Get
				With WordApp
					Return .Name & " " & .Version
				End With
			End Get
		End Property

		''' <summary>プラグインのバージョン</summary>
		''' <returns>○</returns>
		Public ReadOnly Property プラグインのバージョン() As String
			Get
				Return My.Application.Info.Title & " " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision
			End Get
		End Property

		''' <summary>現在開いている文章のファイル名</summary>
		''' <returns>○</returns>
		Public ReadOnly Property ファイル名() As String
			Get
				Return WordApp.ActiveDocument.FullName
			End Get
		End Property

		''' <summary>ワードが起動しているかどうか</summary>
		''' <returns>○</returns>
		Public ReadOnly Property 起動中() As Boolean
			Get
				On Error Resume Next
				If WordApp Is Nothing Then WordApp = CType(GetObject(, "Word.Application"), Word.Application)
				Return Not WordApp Is Nothing
			End Get
		End Property

		''' <summary>ワードが起動しているかどうか</summary>
		Public Function 起動中かどうか() As Boolean
			Return 起動中
		End Function

		''' <summary></summary>
		''' <returns>◎</returns>
		Public Property 警告表示() As Boolean
			Get
				Return WordApp.DisplayAlerts <> Word.WdAlertLevel.wdAlertsNone
			End Get
			Set(ByVal Value As Boolean)
				If Value Then
					WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll
				Else
					WordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
				End If
			End Set
		End Property

		''' <summary>現在開いている文章の一覧</summary>
		''' <returns>○</returns>
		Public ReadOnly Property 文章一覧() As String()
			Get
				Dim Arr As String()
				Dim I As Integer
				Dim TmpDocument As Word.Document
				With WordApp
					ReDim Arr(.Documents.Count - 1)
					For I = 1 To .Documents.Count
						TmpDocument = .Documents.Item(CInt(I))
						Arr(I - 1) = TmpDocument.Name
					Next
					Return Arr
				End With
			End Get
		End Property

		''' <summary></summary>
		''' <returns>○</returns>
		Public ReadOnly Property 選択() As ワード選択範囲
			Get
				Return New ワード選択範囲(Me, WordApp.Selection)
			End Get
		End Property

		''' <summary></summary>
		''' <returns>○</returns>
		Public ReadOnly Property 選択範囲() As ワード選択範囲
			Get
				Return New ワード選択範囲(Me, WordApp.Selection)
			End Get
		End Property

		''' <summary></summary>
		''' <returns>○</returns>
		<設定項目("現在文章", "アクティブドキュメント")>
		Public ReadOnly Property 現在文章() As 文章
			Get
				Dim ActiveDocument As Word.Document = WordApp.ActiveDocument
				If ActiveDocument Is Nothing Then Return Nothing

				Dim Book As New 文章(Me, ActiveDocument)
				Use(Book)
				Return Book
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Word.Application
			Get
				Return WordApp
			End Get
		End Property

#End Region

		Dim DocList As DocumentCollection
#Region "IObjectContainer"
		Public ReadOnly Property 子部品一覧() As Object() Implements IObjectContainer.子部品一覧
			Get
				If DocList Is Nothing Then Return Nothing
				Dim Items(DocList.Count) As Object
				DocList.CopyTo(CType(Items, 文章()))
				Return Items
			End Get
		End Property
		<除外>
		Public Overloads Function TryGeyValue(ByVal Key As String, ByRef Value As IProduireClass) As Boolean Implements IObjectContainer.TryGeyValue
			If DocList Is Nothing Then Return False
			Dim Result As Boolean = DocList.Contains(Key)
			If Result Then Value = DocList(Key)
			Return Result
		End Function
		Public ReadOnly Property Name() As String Implements IObjectContainer.Name
			Get
				Return ""
			End Get
		End Property
#End Region

		Protected Overrides Sub Finalize()
			If Not WordApp Is Nothing Then Marshal.FinalReleaseComObject(WordApp)
			MyBase.Finalize()
		End Sub
	End Class

	Public Class DocumentCollection
		Inherits List(Of 文章)
		Implements IDisposable

		Dim Docs As Word.Documents
		ReadOnly WordApp As ワード

		Public Sub New(ByVal WordApp As ワード, ByVal Docs As Word.Documents)
			Me.WordApp = WordApp
			Me.Docs = Docs
		End Sub

		Default Public Overloads ReadOnly Property Item(ByVal key As String) As IProduireClass
			Get
				On Error Resume Next
				Dim Arr() As String
				Dim I As Integer, Doc As Word.Document

				If IsNumeric(key) Then
					Doc = Docs.Item(CShort(key))
				Else
					Doc = Docs.Item(CStr(key))
				End If
				Return New 文章(WordApp, Doc)
			End Get
		End Property

		Public Overloads Function Contains(ByVal key As String) As Boolean
			Dim Doc As Word.Document
			For Each Doc In Docs
				If Doc.FullName = key Then
					Return True
					Exit For
				End If
			Next
			Return False
		End Function

		Protected Overridable Sub Dispose(ByVal disposing As Boolean)
			If Not Docs Is Nothing Then
				If disposing Then
					Me.Clear()
					Marshal.FinalReleaseComObject(Docs)
					Docs = Nothing
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
			Marshal.FinalReleaseComObject(Docs)
			MyBase.Finalize()
		End Sub
	End Class

	''' <summary>文章</summary>
	<種類(DocUrl:="/office/document.htm")>
	Public Class 文章
		Inherits DisposableObject
		Implements IProduireClass

		Dim WordApp As Word.Application
		Public MyDoc As Word.Document
		ReadOnly RdrWord As ワード

		Dim Bookmarks As Word.Bookmarks

		Public Sub New(ByVal RdrWord As ワード, ByVal MyDoc As Word.Document)
			Me.RdrWord = RdrWord
			Me.WordApp = MyDoc.Application
			Me.MyDoc = MyDoc
			Use(MyDoc)
		End Sub

#Region "手順"
		''' <summary>指定したファイルへ文書を保存します</summary>
		<自分を>
		Public Sub 保存()
			With MyDoc
				.Save()
				'.Saved = True
				'WordApp.Visible = True
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を>
		Public Sub 保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyDoc
				.SaveAs(CType(ファイル名, Object))
				'.Saved = True
				'WordApp.Visible = True
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("PDFで"), 動詞("保存")>
		Public Sub PDFで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyDoc
				.ExportAsFixedFormat(ファイル名, Word.WdExportFormat.wdExportFormatPDF)
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("XPSで"), 動詞("保存")>
		Public Sub XPSで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyDoc
				.ExportAsFixedFormat(ファイル名, Word.WdExportFormat.wdExportFormatXPS)
			End With
		End Sub

		''' <summary>指定した名前の図形を取得します</summary>
		''' <remarks>【名前】の</remarks>
		<名詞手順>
		Public Function 図形(<既定()> ByVal Pos As String) As ワード図形
			On Error Resume Next
			Dim Shapes As Word.Shapes = GetShapes()
			Dim MyShape As Word.Shape = CType(Shapes.Range(CType(Pos, Object)), Word.Shape)
			Use(MyShape)

			If MyShape Is Nothing Then
				Throw New ProduireException("指定した図形'" + Pos + "'は見つかりません。", ワード.ERRORBASE + 2)
				Return Nothing
			End If
			Return New ワード図形(RdrWord, MyShape)
		End Function

		Private Function GetShapes() As Word.Shapes
			Dim Shapes As Word.Shapes = MyDoc.Shapes
			Use(Shapes)
			Return Shapes
		End Function

		''' <summary>文章を閉じます</summary>
		<自分を>
		Public Sub 閉じる()
			MyDoc.Close()
		End Sub

		''' <summary>ワークブックを閉じます</summary>
		<自分を, 補語("保存せずに")>
		Public Sub 保存せずに閉じる()
			MyDoc.Close(False)
		End Sub

		''' <summary>しおりを登録します</summary>
		''' <remarks>【自分】に、【ブックマーク名】という、ブックマークを</remarks>
		<自分へ, 補語("ブックマークを"), 動詞("登録")>
		Public Sub ブックマークを登録(<という()> ByVal Name As String)
			If Bookmarks Is Nothing Then
				Bookmarks = MyDoc.Bookmarks
				Use(Bookmarks)
			End If
			Bookmarks.Add(Name)
		End Sub

		''' <summary>登録したしおりへ移動します</summary>
		''' <remarks>【自分】で【ブックマーク名】という、ブックマークへ</remarks>
		<自分で, 補語("ブックマークへ"), 動詞("移動")>
		Public Sub ブックマークへ移動(<という()> ByVal Name As String)
			Try
				If Bookmarks Is Nothing Then
					Bookmarks = MyDoc.Bookmarks
					Use(Bookmarks)
				End If
				Dim Bookmark As Word.Bookmark
				Bookmark = Bookmarks.Item(CStr(Name))
				Use(Bookmark)
				Bookmark.Select()
			Catch
				Throw New ProduireException(Err.Description, Err.Number)
			End Try
		End Sub

		''' <summary>文章を印刷します</summary>
		<自分を>
		Public Sub 印刷()
			MyDoc.PrintOut()
		End Sub

		''' <summary>文章の印刷見本を表示します
		''' 再度実行すると閉じます</summary>
		<自分を>
		Public Sub プレビュー()
			WordApp.PrintPreview = Not MyDoc.Application.PrintPreview
		End Sub

		''' <summary>文章に含まれるマクロを実行します</summary>
		'''　<remarks>【マクロ名】を〈【引数】で〉</remarks>
		<自分("にある")>
		Public Function マクロ実行(<を> ByVal マクロ名 As String, <で()> ByVal 引数 As String()) As Object
			Dim Arr As New List(Of Object)
			Dim Result As Object
			Arr.Add(マクロ名)
			If Not 引数 Is Nothing Then Arr.AddRange(引数)
			Result = RunMacro(Arr.ToArray())
			If Err.Number <> 0 Then Throw New ProduireException(Err.Description, ワード.ERRORBASE + 7) : Err.Clear()
			Return Result
		End Function

#End Region

#Region "設定項目"

		''' <summary>現在開いている文章のファイル名</summary>
		''' <returns>□</returns>
		Public ReadOnly Property ファイル名() As String
			Get
				Return MyDoc.FullName
			End Get
		End Property


		''' <summary>現在開いている文章の本文</summary>
		''' <returns>□</returns>
		Public Property 内容() As String
			Get
				Return MyDoc.Content.Text
			End Get
			Set(ByVal Value As String)
				MyDoc.Content.Text = Value
			End Set
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As Word.Document
			Get
				Return MyDoc
			End Get
		End Property

#End Region

#Region "図形関連"

		''' <summary>シートへ画像を追加します</summary>
		<自分へ, 手順名("画像として", "追加")>
		Public Function 画像として追加する(<を> ByVal Value As String) As ワード図形
			Dim Shapes As Word.Shapes = GetShapes()
			Dim MyShape As Word.Shape = Shapes.AddPicture(Value, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, -1, -1)
			Use(MyShape)
			Return New ワード図形(RdrWord, MyShape)
		End Function

		''' <summary>スライドにある図形の一覧</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 図形一覧() As String()
			Get
				Dim I As Integer
				Dim Shapes As Word.Shapes = GetShapes()
				Dim TmpShape As Word.Shape

				Dim Result As String()
				ReDim Result(Shapes.Count - 1)
				For I = 1 To Shapes.Count
					TmpShape = Shapes.Item(CInt(I))
					Result(I - 1) = TmpShape.Name
					Use(TmpShape)
				Next
				Return Result
			End Get
		End Property

#End Region

		Private Function RunMacro(ByVal oRunArgs As Object()) As Object
			Return GetType(Word.Application).InvokeMember("Run", BindingFlags.Default Or BindingFlags.InvokeMethod, Nothing, WordApp, oRunArgs)
		End Function

		Protected Overrides Sub Finalize()
			Marshal.FinalReleaseComObject(MyDoc)
			MyBase.Finalize()
		End Sub
	End Class
End Namespace
