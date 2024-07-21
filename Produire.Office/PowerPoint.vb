' プロデル Microsoft Office日本語プログラミングライブラリ
' Copyright(C) 2007-2024 irelang.jp https://github.com/utopiat-ire/
Option Strict On
Imports System.Drawing
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports MsoTriState = Microsoft.Office.Core.MsoTriState
Imports utopiat
Imports System.Reflection

Namespace パワーポイント

	''' <summary>
	''' Microsoft PowerPointを操作する種類です。
	''' パワーポイントは、発表資料とスライドで構成されています。
	''' </summary>
	<種類(DocUrl:="/office/powerpoint.htm")>
	Public Class パワーポイント
		Inherits DisposableObject
		Implements IProduireStaticClass
		Implements IObjectContainer

		Public Const ERRORBASE As Integer = 2200
		Dim PptApp As PowerPoint.Application

#Region "手順"
		''' <summary>PowerPointを起動します</summary>
		<自分を> Public Sub 起動()
			GetPptApp()
			PptApp.Visible = MsoTriState.msoTrue
			Tips.SetForegroundWindow(CType(PptApp.HWND, IntPtr))
		End Sub

		Private Sub GetPptApp()
			On Error Resume Next
			Dim IsNotRunning As Boolean
			PptApp = CType(GetObject(, "PowerPoint.Application"), PowerPoint.Application)
			IsNotRunning = (Err().Number <> 0)
			Err.Clear()
			If IsNotRunning Then
				PptApp = New PowerPoint.Application()
				Use(PptApp)
			End If

			Init()
		End Sub

		Private Sub Init()
			PptList = New PresentationCollection(Me, PptApp.Presentations)
			Use(PptList)
		End Sub

		''' <summary>パワーポイントを表示します。</summary>
		<自分を> Public Sub 表示()
			If PptApp Is Nothing Then Throw New ProduireException("エクセルを起動またはリンクしていないためリンクできません。", ERRORBASE + 2)
			PptApp.Visible = MsoTriState.msoTrue
		End Sub

		''' <summary>パワーポイントを非表示にします</summary>
		<自分を> Public Sub 非表示()
			If PptApp Is Nothing Then Throw New ProduireException("エクセルを起動またはリンクしていないためリンクできません。", ERRORBASE + 2)
			PptApp.Visible = MsoTriState.msoFalse
		End Sub

		''' <summary>すでに起動しているパワーポイントとリンクして、プロデルから利用できるようにします。</summary>
		<自分を> Public Sub リンク()
			Try
				PptApp = CType(GetObject(, "PowerPoint.Application"), PowerPoint.Application)
				Use(PptApp)
			Catch When Err.Number = 429
				Throw New ProduireException("エクセルが起動していないためリンクできません。", ERRORBASE + 2)
			End Try

			Init()
		End Sub

		''' <summary>Excelを起動して新しいブックを開きます</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Function 開く() As 発表資料
			起動()
			Dim NewPresent As 発表資料 = 新規発表資料()
			Dim NewSlide As スライド = スライドを追加()
			NewSlide.レイアウト = PowerPoint.PpSlideLayout.ppLayoutTitle
			Return NewPresent
		End Function

		''' <summary>発表資料(プレゼンテーション)を開きます</summary>
		'''　<remarks>【ファイル名】を</remarks>
		<自分で>
		Public Function 開く(<を> ByVal ファイル名 As String) As 発表資料
			On Error Resume Next
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください", ERRORBASE + 3)
			End If

			GetPptApp()

			Dim Current As PowerPoint.Presentation
			With PptApp
				Dim FullPath As String
				If Mid(ファイル名, 2, 1) <> ":" Then
					FullPath = FileUtils.YenSuffix(CurDir()) & ファイル名
				Else
					FullPath = ファイル名
				End If
				Err.Clear()

				Current = PptList.Presentations.Open(FullPath)
				PptApp.Activate()
				If Err.Number <> 0 Then
					Throw New ProduireException(Err.Description, ERRORBASE + 3)
				Else
					PptApp.Visible = MsoTriState.msoCTrue
				End If
			End With
			Dim NewPresent As 発表資料
			NewPresent = New 発表資料(Me, Current)
			PptList.Add(NewPresent)
			Return NewPresent
		End Function

		''' <summary>現在の発表資料を保存します</summary>
		<自分を>
		Public Sub 保存()
			現在発表資料.保存()
		End Sub

		''' <summary>現在の発表資料を保存します</summary>
		'''　<remarks>【ファイル名】へ</remarks>
		<自分を> Public Sub 保存(<へ> ByVal ファイル名 As String)
			現在発表資料.保存(ファイル名)
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("PDFで"), 動詞("保存")>
		Public Sub PDFで保存(<へ> ByVal ファイル名 As String)
			現在発表資料.PDFで保存(ファイル名)
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("XPSで"), 動詞("保存")>
		Public Sub XPSで保存(<へ> ByVal ファイル名 As String)
			現在発表資料.XPSで保存する(ファイル名)
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 動詞("エクスポート")>
		Public Sub エクスポート(<へ> ByVal 出力先 As String, <として()> ByVal 形式 As String, <で(), 省略()> ByVal 大きさ As Size)
			現在発表資料.エクスポート(出力先, 形式, 大きさ)
		End Sub

		''' <summary>新しい発表資料を作成します</summary>
		'''　<remarks></remarks>
		<自分へ>
		Public Function 新規発表資料() As 発表資料
			Dim NewPresent As 発表資料
			NewPresent = New 発表資料(Me, PptApp.Presentations.Add())
			PptList.Add(NewPresent)
			Return NewPresent
		End Function

		''' <summary>スライドまたは発表資料を追加します</summary>
		'''　<remarks>スライドを</remarks>
		<自分へ, 補語("スライドを"), 動詞("追加", "作成")>
		Public Function スライドを追加() As スライド
			Return 現在発表資料.追加()
		End Function

		''' <summary>スライドまたは発表資料を追加します</summary>
		'''　<remarks>発表資料を</remarks>
		<自分へ, 補語("発表資料を"), 動詞("追加", "作成")>
		Public Function 発表資料を追加() As 発表資料
			Return 新規発表資料()
		End Function

		''' <summary>発表資料を追加します</summary>
		'''　<remarks>プレゼンテーションを</remarks>
		<自分へ, 補語("プレゼンテーションを"), 動詞("追加", "作成")>
		Public Function プレゼンテーションを追加() As 発表資料
			Return 新規発表資料()
		End Function

		''' <summary>保存確認せずにPowerPointを終了します</summary>
		<自分を>
		Public Sub 強制終了()
			PptApp.ActivePresentation.Saved = MsoTriState.msoTrue
			終了()
		End Sub

		''' <summary>PowerPointを終了します</summary>
		<自分を> Public Sub 終了()
			PptApp.Quit()
			解放()
		End Sub

		<自分を> Private Sub 解放()
			PptList.Clear()
			PptList.Dispose()
			GC.Collect()
			Dispose()
		End Sub

		''' <summary>発表資料に含まれるマクロを実行します</summary>
		'''　<remarks>【マクロ名】を〈【引数】で〉</remarks>
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

		''' <summary>発表資料を印刷します</summary>
		<自分を> Public Sub 印刷()
			Dim MyPresentation As PowerPoint.Presentation = TryCast(PptApp.ActivePresentation, PowerPoint.Presentation)
			Use(MyPresentation)
			MyPresentation.PrintOut()
		End Sub

		''' <summary>スライドショーを開始します</summary>
		<自分を> Public Sub 開始()
			Dim MyPresentation As PowerPoint.Presentation = TryCast(PptApp.ActivePresentation, PowerPoint.Presentation)
			Use(MyPresentation)
			Dim SlideShowSettings As PowerPoint.SlideShowSettings = MyPresentation.SlideShowSettings
			Use(SlideShowSettings)
			SlideShowSettings.Run()
		End Sub

		''' <summary>次のスライドを進めます</summary>
		<自分を> Public Sub 進ませる()
			Dim MyPresentation As PowerPoint.Presentation = TryCast(PptApp.ActivePresentation, PowerPoint.Presentation)
			Use(MyPresentation)
			Dim Window As PowerPoint.SlideShowWindow = MyPresentation.SlideShowWindow
			Use(Window)
			Dim View As PowerPoint.SlideShowView = Window.View
			Use(View)

			View.Next()
		End Sub

		''' <summary>前のスライドへ戻します</summary>
		<自分を> Public Sub 戻す()
			Dim MyPresentation As PowerPoint.Presentation = TryCast(PptApp.ActivePresentation, PowerPoint.Presentation)
			Use(MyPresentation)
			Dim Window As PowerPoint.SlideShowWindow = MyPresentation.SlideShowWindow
			Use(Window)
			Dim View As PowerPoint.SlideShowView = Window.View
			Use(View)

			View.Previous()
		End Sub

#End Region

#Region "設定項目"
		''' <summary>パワーポイントのバージョン情報を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property バージョン() As String
			Get
				With PptApp
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

		''' <summary>現在選択している発表資料のファイル名を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property ファイル名() As String
			Get
				Return PptApp.ActivePresentation.FullName
			End Get
		End Property

		''' <summary>パワーポイントが現在起動しているかどうかを表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 起動中() As Boolean
			Get
				On Error Resume Next
				If PptApp Is Nothing Then PptApp = CType(GetObject(, "PowerPoint.Application"), PowerPoint.Application)
				Return (Not PptApp Is Nothing)
			End Get
		End Property

		''' <summary>パワーポイントが現在起動しているかどうかを表します</summary>
		Public Function 起動中かどうか() As Boolean
			Return 起動中
		End Function

		''' <summary></summary>
		''' <returns>◎</returns>
		Public Property 警告表示() As Boolean
			Get
				Return PptApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll
			End Get
			Set(ByVal Value As Boolean)
				If Value Then
					PptApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsAll
				Else
					PptApp.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone
				End If
			End Set
		End Property

		''' <summary>現在開いている発表資料の一覧を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 発表資料一覧() As String()
			Get
				If Not PptApp Is Nothing Then
					Dim Arr As String()
					With PptApp
						Dim I As Integer
						Dim TmpPresentation As PowerPoint.Presentation
						Dim TmpPresentations As PowerPoint.Presentations = PptApp.Presentations
						Use(TmpPresentations)
						ReDim Arr(TmpPresentations.Count - 1)
						For I = 1 To TmpPresentations.Count
							TmpPresentation = TmpPresentations.Item(I)
							Arr(I - 1) = TmpPresentation.Name
							Use(TmpPresentation)
						Next
					End With
					Return Arr
				Else
					Return Nothing
				End If
			End Get
		End Property

		''' <summary>現在選択している発表資料を表します</summary>
		''' <returns>□</returns>
		Public Property 発表資料() As String
			Get
				Dim Presentation As PowerPoint.Presentation
				Presentation = PptApp.ActivePresentation
				Return Presentation.Name
			End Get
			Set(ByVal value As String)
				If value.Length = 0 Then Exit Property
				Dim Presentations As PowerPoint.Presentations, Presentation As PowerPoint.Presentation
				Presentations = PptApp.Presentations
				Presentation = Presentations.Item(CStr(value))
				'Book.Activate()
			End Set
		End Property

		''' <summary>現在選択しているスライドを表します</summary>
		''' <returns>□</returns>
		Public Property スライド() As Integer
			Get
				Dim Document As PowerPoint.DocumentWindow = PptApp.ActiveWindow
				Use(Document)
				Dim View As PowerPoint.View = Document.View
				Use(View)
				Dim Slide As PowerPoint.Slide = TryCast(Document.View.Slide, PowerPoint.Slide)
				Use(Slide)
				Return Slide.SlideNumber
			End Get

			Set(ByVal value As Integer)
				Dim Presentation As PowerPoint.Presentation = PptApp.ActivePresentation
				Use(Presentation)
				Dim Slides As PowerPoint.Slides = Presentation.Slides
				Use(Slides)
				Dim Slide As PowerPoint.Slide = Slides.Item(value)
				Use(Slide)
				Slide.Select()
			End Set
		End Property

		''' <summary>現在開いている発表資料の数を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 発表資料数() As Integer
			Get
				Return PptList.Count
			End Get
		End Property

		''' <summary>パワーポイントが画面に表示されているかどうかを表します</summary>
		Public Property 表示状態() As Boolean
			Get
				If PptApp Is Nothing Then Return False
				Return (PptApp.Visible = MsoTriState.msoTrue)
			End Get
			Set(ByVal value As Boolean)
				If PptApp Is Nothing Then Throw New ProduireException("エクセルを起動またはリンクしていないためリンクできません。", ERRORBASE + 2) : Return
				PptApp.Visible = DirectCast(IIf(value, MsoTriState.msoTrue, MsoTriState.msoFalse), MsoTriState)
			End Set
		End Property

		''' <summary></summary>
		''' <returns>○</returns>
		Public ReadOnly Property 現在発表資料() As 発表資料
			Get
				Dim Present As New 発表資料(Me, PptApp.ActivePresentation)
				Return Present
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As PowerPoint.Application
			Get
				Return PptApp
			End Get
		End Property

#End Region

#Region "IObjectContainer"
		Dim PptList As PresentationCollection
		Public ReadOnly Property 子部品一覧() As Object() Implements IObjectContainer.子部品一覧
			Get
				If PptList Is Nothing Then Return Nothing
				Dim Items(PptList.Count) As Object
				PptList.CopyTo(CType(Items, 発表資料()))
				Return Items
			End Get
		End Property
		<除外>
		Public Overloads Function TryGeyValue(ByVal Key As String, ByRef Value As IProduireClass) As Boolean Implements IObjectContainer.TryGeyValue
			If PptList Is Nothing Then Return False
			Dim Result As Boolean = PptList.Contains(Key)
			If Result Then Value = PptList(Key)
			Return Result
		End Function
		Public ReadOnly Property Name() As String Implements IObjectContainer.Name
			Get
				Return ""
			End Get
		End Property
#End Region

		Friend Function pGetSlide(ByRef SheetName As String) As PowerPoint.Slide
			If IsNumeric(SheetName) Then
				pGetSlide = CType(PptList.Item(CShort(SheetName)), PowerPoint.Slide)
			Else
				pGetSlide = CType(PptList.Item(SheetName), PowerPoint.Slide)
			End If
		End Function

		Private Function RunMacro(ByVal oRunArgs As Object()) As Object
			Return GetType(PowerPoint.Application).InvokeMember("Run", BindingFlags.Default Or BindingFlags.InvokeMethod, Nothing, PptApp, oRunArgs)
		End Function

	End Class

	''' <summary>発表資料</summary>
	<種類(DocUrl:="/office/presentation.htm")>
	Public Class 発表資料
		Inherits DisposableObject
		Implements IProduireClass

		Dim PptApp As PowerPoint.Application
		Public MyPresentation As PowerPoint.Presentation
		ReadOnly RdrPpt As パワーポイント

		Public Sub New(ByVal RdrPpt As パワーポイント, ByVal MyPresentation As PowerPoint.Presentation)
			PptApp = MyPresentation.Application
			Me.MyPresentation = MyPresentation
			Me.RdrPpt = RdrPpt
			Use(MyPresentation)
		End Sub

#Region "手順"
		''' <summary>指定したスライドを選択します</summary>
		'''　<remarks></remarks>
		<自分を>
		Public Function 選択(<を> ByVal シート名 As String) As スライド
			Dim Slide As PowerPoint.Slide
			Slide = MyPresentation.Slides.Item(シート名)
			Use(Slide)
			Slide.Select()
			Return New スライド(RdrPpt, Slide)
		End Function

		''' <summary>スライドを追加します</summary>
		'''　<remarks></remarks>
		<自分へ, 動詞("追加", "作成")>
		Public Function 追加(<へ> ByVal 番号 As Integer) As スライド
			Dim NewSlide As PowerPoint.Slide
			Dim Slides As PowerPoint.Slides = MyPresentation.Slides
			Use(Slides)
			NewSlide = Slides.Add(番号, PowerPoint.PpSlideLayout.ppLayoutBlank)
			Use(NewSlide)
			NewSlide.Select()
			Return New スライド(RdrPpt, NewSlide)
		End Function

		''' <summary>スライドを追加します</summary>
		'''　<remarks></remarks>
		<自分へ, 動詞("追加", "作成")>
		Public Function 追加() As スライド
			Dim NewSlide As PowerPoint.Slide
			Dim Slides As PowerPoint.Slides = MyPresentation.Slides
			Use(Slides)
			NewSlide = Slides.Add(Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank)
			Use(NewSlide)
			NewSlide.Select()
			Return New スライド(RdrPpt, NewSlide)
		End Function

		''' <summary>現在の発表資料を保存します</summary>
		<自分を> Public Sub 保存()
			'On Error Resume Next
			With MyPresentation
				.Save()
				'.Saved = MsoTriState.msoTrue
			End With
		End Sub

		''' <summary>現在の発表資料を保存します</summary>
		'''　<remarks>【ファイル名】へ</remarks>
		<自分を> Public Sub 保存(<へ> ByVal ファイル名 As String)
			'On Error Resume Next
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyPresentation
				.SaveAs(ファイル名)
				'.Saved = MsoTriState.msoTrue
				'.Application.Visible = MsoTriState.msoTrue
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("PDFで"), 動詞("保存")>
		Public Sub PDFで保存(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyPresentation
				.ExportAsFixedFormat(ファイル名, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF)
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 補語("XPSで"), 動詞("保存")>
		Public Sub XPSで保存する(<へ> ByVal ファイル名 As String)
			If Len(ファイル名) = 0 Then
				Throw New ProduireException("ファイル名を指定してください。")
			End If
			With MyPresentation
				.ExportAsFixedFormat(ファイル名, PowerPoint.PpFixedFormatType.ppFixedFormatTypeXPS)
			End With
		End Sub

		''' <summary>指定したファイルへ文書を保存します</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 動詞("エクスポート")>
		Public Sub エクスポート(<へ> ByVal 出力先 As String, <として()> ByVal 形式 As String, <で()> ByVal 大きさ As Size)
			If Len(出力先) = 0 Then
				Throw New ProduireException("出力先を指定してください。")
			End If
			With MyPresentation
				.Export(出力先, 形式, 大きさ.Width, 大きさ.Height)
			End With
		End Sub

		''' <summary>発表資料を閉じます</summary>
		<自分を> Public Sub 閉じる()
			MyPresentation.Close()
		End Sub

		''' <summary>発表資料を印刷します</summary>
		<自分を>
		Public Sub 印刷()
			MyPresentation.PrintOut()
		End Sub

		''' <summary>発表資料に含まれるマクロを実行します</summary>
		''' <remarks>【マクロ名】を〈【引数】で〉</remarks>
		<自分("にある")>
		Public Function マクロ実行(<を> ByVal マクロ名 As String, <で(), 省略()> ByVal 引数 As String()) As Object
			Dim Arr As New List(Of Object)
			Dim Result As Object
			Arr.Add(マクロ名)
			If Not 引数 Is Nothing Then Arr.AddRange(引数)
			Result = RunMacro(Arr.ToArray())
			If Err.Number <> 0 Then Throw New ProduireException(Err.Description, パワーポイント.ERRORBASE + 7) : Err.Clear()
			Return Result
		End Function

		Private Function RunMacro(ByVal oRunArgs As Object()) As Object
			Return GetType(PowerPoint.Application).InvokeMember("Run", BindingFlags.Default Or BindingFlags.InvokeMethod, Nothing, PptApp, oRunArgs)
		End Function
#End Region

#Region "設定項目"
		''' <summary>ファイル名を表します</summary>
		''' <returns>□</returns>
		Public ReadOnly Property ファイル名() As String
			Get
				Return MyPresentation.FullName
			End Get
		End Property

		''' <summary>スライドの番号を表します</summary>
		''' <returns>◎</returns>
		Public Property スライド番号() As Integer
			Get
				Dim Document As PowerPoint.DocumentWindow = PptApp.ActiveWindow
				Use(Document)
				Dim View As PowerPoint.View = Document.View
				Use(View)
				Dim Slide As PowerPoint.Slide = TryCast(Document.View.Slide, PowerPoint.Slide)
				Use(Slide)
				Return Slide.SlideNumber
			End Get
			Set(ByVal value As Integer)
				Dim Slides As PowerPoint.Slides, Slide As PowerPoint.Slide
				Slides = MyPresentation.Slides
				Slide = Slides.Item(value)
				Slide.Select()
			End Set
		End Property

		''' <summary>現在スライドを表します</summary>
		''' <returns>◎</returns>
		Public Property 現在スライド() As スライド
			Get
				Dim Window As PowerPoint.DocumentWindow = PptApp.ActiveWindow
				Use(Window)
				Dim View As PowerPoint.View = Window.View
				Use(View)
				Dim Slide As PowerPoint.Slide = TryCast(Window.View.Slide, PowerPoint.Slide)
				Use(Slide)
				Return New スライド(RdrPpt, Slide)
			End Get
			Set(ByVal value As スライド)
				value.MySlide.Select()
			End Set
		End Property

		''' <summary>スライドの一覧</summary>
		''' <returns>□</returns>
		Public ReadOnly Property スライド一覧() As String()
			Get
				Dim Result As String()
				With MyPresentation
					Dim TmpSlide As PowerPoint.Slide
					ReDim Result(.Slides.Count - 1)
					Dim I As Integer
					For I = 1 To .Slides.Count
						TmpSlide = .Slides.Item(I)
						Result(I - 1) = TmpSlide.Name
					Next
				End With
				Return Result
			End Get
		End Property

		''' <summary>スライドの総数</summary>
		''' <returns>□</returns>
		Public ReadOnly Property スライド数() As Integer
			Get
				Return MyPresentation.Slides.Count
			End Get
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As PowerPoint.Presentation
			Get
				Return MyPresentation
			End Get
		End Property

#End Region

	End Class

	''' <summary>スライド</summary>
	<種類(DocUrl:="/office/slide.htm")>
	Public Class スライド
		Inherits DisposableObject
		Implements IProduireClass

		ReadOnly PptApp As PowerPoint.Application
		Friend MySlide As PowerPoint.Slide
		Dim RdrPpt As パワーポイント

		Public Sub New(ByVal MainClass As パワーポイント, ByVal MySheet As PowerPoint.Slide)
			Me.RdrPpt = MainClass
			Me.MySlide = MySheet
			Me.PptApp = MySheet.Application
			Use(MySheet)
		End Sub

#Region "手順"

		''' <summary>スライドを選択します</summary>
		<自分を> Public Sub 選択する()
			On Error Resume Next
			MySlide.Select()
		End Sub

		''' <summary>スライドを消します</summary>
		<自分を> Public Sub 消す()
			MySlide.Delete()
		End Sub

		''' <summary>スライドをクリップボードへコピーします</summary>
		<自分を> Public Sub コピー()
			MySlide.Copy()
		End Sub

		''' <summary>スライドを移動します</summary>
		'''　<remarks>《移動先》へ</remarks>
		<自分を> Public Sub 移動(<へ> ByVal 移動先 As String)
			If Len(移動先) <> 0 Then
				RdrPpt.pGetSlide(移動先).Select()
			End If
		End Sub

		''' <summary>スライドを画像としてエクスポートします</summary>
		''' <remarks>【ファイル名】へ</remarks>
		<自分を, 動詞("エクスポート")>
		Public Sub エクスポート(<へ> ByVal 出力ファイル名 As String, <として()> ByVal 形式 As String, <で(), 省略()> ByVal 大きさ As Size)
			If Len(出力ファイル名) = 0 Then
				Throw New ProduireException("出力先を指定してください。")
			End If
			With MySlide
				.Export(出力ファイル名, 形式, 大きさ.Width, 大きさ.Height)
			End With
		End Sub

#End Region

#Region "設定項目"
		''' <summary>スライドの名前</summary>
		Public Property 名前() As String
			Get
				Return MySlide.Name
			End Get
			Set(ByVal Value As String)
				MySlide.Name = Value
			End Set
		End Property

		''' <summary>スライドにある図形の一覧</summary>
		''' <returns>□</returns>
		Public ReadOnly Property 図形一覧() As String()
			Get
				Dim I As Integer
				Dim Shapes As PowerPoint.Shapes = GetShapes()
				Dim TmpShape As PowerPoint.Shape

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

		''' <summary>スライドのレイアウト</summary>
		Public Property レイアウト() As PowerPoint.PpSlideLayout
			Get
				Return MySlide.Layout

			End Get
			Set(ByVal Value As PowerPoint.PpSlideLayout)
				MySlide.Layout = Value
			End Set
		End Property

		''' <summary></summary>
		''' <returns>□</returns>
		<設定項目("元実体")>
		Public ReadOnly Property 元実体() As PowerPoint.Slide
			Get
				Return MySlide
			End Get
		End Property


#End Region

#Region "図形関連"

		''' <summary>シートへ画像を追加します</summary>
		<自分へ, 手順名("画像として", "追加")>
		Public Function 画像として追加する(<を> ByVal Value As String) As パワーポイント図形
			Dim Shapes As PowerPoint.Shapes = GetShapes()
			Dim MyShape As PowerPoint.Shape = Shapes.AddPicture(Value, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0, -1, -1)
			Use(MyShape)
			Return New パワーポイント図形(RdrPpt, MyShape)
		End Function

		''' <summary>指定した名前の図形を取得します</summary>
		''' <remarks>【名前】の</remarks>
		<名詞手順>
		Public Function 図形(<既定()> ByVal Pos As String) As パワーポイント図形
			On Error Resume Next
			Dim Shapes As PowerPoint.Shapes = GetShapes()
			Dim MyShape As PowerPoint.Shape = CType(Shapes.Range(CType(Pos, Object)), PowerPoint.Shape)
			Use(MyShape)

			If MyShape Is Nothing Then
				Throw New ProduireException("指定した図形'" + Pos + "'は見つかりません。", パワーポイント.ERRORBASE + 2)
				Return Nothing
			End If
			Return New パワーポイント図形(RdrPpt, MyShape)
		End Function

		Private Function GetShapes() As PowerPoint.Shapes
			Dim Shapes As PowerPoint.Shapes = MySlide.Shapes
			Use(Shapes)
			Return Shapes
		End Function

		Friend Function GetShape(ByRef ShapeName As String) As PowerPoint.Shape
			Dim Shapes As PowerPoint.Shapes = GetShapes()
			Return Shapes.Item(CStr(ShapeName))
		End Function

#End Region

	End Class

	Public Class PresentationCollection
		Inherits List(Of 発表資料)
		Implements IDisposable

		ReadOnly PptApp As パワーポイント
		Public Presentations As PowerPoint.Presentations

		Public Sub New(ByVal PptApp As パワーポイント, ByVal Presentation As PowerPoint.Presentations)
			Me.PptApp = PptApp
			Me.Presentations = Presentation
		End Sub

		Public Overloads Function Contains(ByVal key As String) As Boolean
			Dim Presentation As PowerPoint.Presentation
			For Each Presentation In Presentations
				If Presentation.FullName = key Then
					Return True
					Exit For
				End If
			Next
			Return False
		End Function

		Public Overloads Sub Clear()
			Dim MyPresentation As 発表資料
			For Each MyPresentation In Me
				MyPresentation.Dispose()
			Next
			MyBase.Clear()
		End Sub

		Public Sub Refresh()
			Dim Presentation As PowerPoint.Presentation
			For Each Presentation In Presentations
				Me.Add(New 発表資料(PptApp, Presentation))
			Next
		End Sub

		Protected Overridable Sub Dispose(ByVal disposing As Boolean)
			If Not Presentations Is Nothing Then
				If disposing Then
					Me.Clear()
					Marshal.FinalReleaseComObject(Presentations)
					Presentations = Nothing
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

		Default Public Overloads ReadOnly Property Item(ByVal key As String) As 発表資料
			Get
				On Error Resume Next
				Dim Arr() As String
				Dim I As Integer, Present As PowerPoint.Presentation

				If IsNumeric(key) Then
					Present = Presentations.Item(CShort(key))
				Else
					Present = Presentations.Item(key)
				End If
				Return New 発表資料(PptApp, Present)
			End Get
		End Property
	End Class

End Namespace
