' vba-compat.vbs - VBA固有の機能をVBSでモックする互換レイヤー
' 注意: GetScriptDir()はrun-tests.vbsで動的に定義される（パスを埋め込むため）

' ============================================
' Debug オブジェクトのモック
' ============================================
' VBSでは Print が予約語のため、Debug.Print は直接モックできない
' 代わりに変換スクリプトで Debug.Print を DebugPrint に変換する
Sub DebugPrint(ByVal msg)
    WScript.Echo msg
End Sub

' ============================================
' VBA関数のモック
' ============================================

' FreeFile - VBSでは使えないが、モックとして0を返す
Function FreeFile()
    FreeFile = 0
End Function

' DoEvents - VBSでは何もしない
Sub DoEvents()
    ' No-op
End Sub

' IsMissing - VBSでは常にFalseを返す（Optionalパラメータは変換時に処理済み）
Function IsMissing(ByVal arg)
    IsMissing = IsEmpty(arg)
End Function

' ============================================
' Collection クラス (VBSにはないのでモック)
' 注意: VBSのカスタムClassはFor Eachに対応していない
' ============================================
Class Collection
    Private items_

    Private Sub Class_Initialize()
        Set items_ = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub Add(ByVal item, ByVal key)
        If key = "" Then
            key = "item_" & (items_.Count + 1)
        End If
        items_.Add key, item
    End Sub

    Public Sub Remove(ByVal index)
        If IsNumeric(index) Then
            Dim i, k
            i = 0
            For Each k In items_.Keys
                i = i + 1
                If i = index Then
                    items_.Remove k
                    Exit Sub
                End If
            Next
        Else
            items_.Remove index
        End If
    End Sub

    Public Property Get Item(ByVal index)
        If IsNumeric(index) Then
            Dim i, k
            i = 0
            For Each k In items_.Keys
                i = i + 1
                If i = index Then
                    If IsObject(items_(k)) Then
                        Set Item = items_(k)
                    Else
                        Item = items_(k)
                    End If
                    Exit Property
                End If
            Next
        Else
            If IsObject(items_(index)) Then
                Set Item = items_(index)
            Else
                Item = items_(index)
            End If
        End If
    End Property

    Public Property Get Count()
        Count = items_.Count
    End Property

    Public Sub Clear()
        items_.RemoveAll
    End Sub

    ' 内部Dictionaryを取得（For Each用）
    Public Function GetDict()
        Set GetDict = items_
    End Function
End Class

' Collection のファクトリ関数
Function CreateCollection()
    Set CreateCollection = New Collection
End Function

' ============================================
' FileSystemObject ヘルパー
' ============================================
Dim fso_compat
Set fso_compat = CreateObject("Scripting.FileSystemObject")

' VBA互換のファイル書き込み
Sub WriteTextFile(ByVal path, ByVal text)
    Dim f
    Set f = fso_compat.CreateTextFile(path, True)
    f.Write text
    f.Close
End Sub

' VBA互換のファイル読み込み
Function ReadTextFile(ByVal path)
    If fso_compat.FileExists(path) Then
        Dim f
        Set f = fso_compat.OpenTextFile(path, 1)
        If f.AtEndOfStream Then
            ReadTextFile = ""
        Else
            ReadTextFile = f.ReadAll
        End If
        f.Close
    Else
        ReadTextFile = ""
    End If
End Function

' ファイル存在チェック
Function FileExists(ByVal path)
    FileExists = fso_compat.FileExists(path)
End Function

' フォルダ存在チェック
Function FolderExists(ByVal path)
    FolderExists = fso_compat.FolderExists(path)
End Function

' ============================================
' 型変換関数
' ============================================
' CLng, CStr, CInt, CDbl, CBool, CDate は VBS標準で使える

' CLngPtr - VBAのポインタ型、VBSではCLngにフォールバック
Function CLngPtr(ByVal value)
    CLngPtr = CLng(value)
End Function

' CVErr - VBAのエラー値、VBSでは0を返す
Function CVErr(ByVal errNum)
    CVErr = 0
End Function

' ============================================
' 配列関数
' ============================================
' LBound, UBound, Array, Split, Join は VBS標準で使える

' ReDim Preserve の代替（VBSでも使えるが念のため）
' 注意: VBSのReDim Preserveは最後の次元のみ変更可能

' ============================================
' Utils モジュールのモック
' VBA: Utils.Fail → VBS: UtilsFail
' ============================================
Sub UtilsFail(ByVal code, ByVal message)
    Err.Raise code, "VBA_TEST", message
End Sub

Sub UtilsWriteTextFile(ByVal path, ByVal text)
    WriteTextFile path, text
End Sub

' ============================================
' テスト用アサーション関数
' ============================================
Sub Assert(ByVal condition, ByVal message)
    If Not condition Then
        Err.Raise 9999, "Assert", "Assertion failed: " & message
    End If
End Sub

Sub AssertEqual(ByVal expected, ByVal actual, ByVal message)
    If expected <> actual Then
        Err.Raise 9999, "AssertEqual", message & " - Expected: " & CStr(expected) & ", Actual: " & CStr(actual)
    End If
End Sub

Sub AssertTrue(ByVal condition, ByVal message)
    Assert condition, message
End Sub

Sub AssertFalse(ByVal condition, ByVal message)
    Assert Not condition, message
End Sub

Sub AssertNothing(ByVal obj, ByVal message)
    If Not obj Is Nothing Then
        Err.Raise 9999, "AssertNothing", message & " - Object is not Nothing"
    End If
End Sub

Sub AssertNotNothing(ByVal obj, ByVal message)
    If obj Is Nothing Then
        Err.Raise 9999, "AssertNotNothing", message & " - Object is Nothing"
    End If
End Sub

' ============================================
' Excel/Workbook関連のモック（必要に応じて拡張）
' ============================================
' ThisWorkbook.Path → GetScriptDir() に変換（変換スクリプトで処理）
' ActiveWorkbook, Worksheets等は未サポート（必要に応じて追加）

' ============================================
' その他のVBA互換関数
' ============================================

' Environ - 環境変数を取得
Function Environ(ByVal name)
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    Environ = shell.ExpandEnvironmentStrings("%" & name & "%")
    If Environ = "%" & name & "%" Then
        Environ = ""
    End If
End Function

' Timer - 午前0時からの経過秒数（小数点以下含む）
Function Timer()
    Timer = CDbl(Hour(Now) * 3600 + Minute(Now) * 60 + Second(Now))
End Function

' Sleep - 指定ミリ秒待機
Sub Sleep(ByVal milliseconds)
    WScript.Sleep milliseconds
End Sub

' ============================================
' 追加のVBA関数モック
' ============================================

' Format - VBSでも基本的に使えるが、念のため
' VBSのFormatDateTimeやFormatNumberを使う

' MsgBox - VBSでも使えるが、CIではダミーを返す
' 注意: CI環境でダイアログを出さないようにするため、常に vbOK を返す
Function MsgBoxCI(ByVal prompt, ByVal buttons, ByVal title)
    DebugPrint "[MsgBox] " & title & ": " & prompt
    MsgBoxCI = 1 ' vbOK
End Function

' InputBox - CI環境ではダミー値を返す
Function InputBoxCI(ByVal prompt, ByVal title, ByVal default)
    DebugPrint "[InputBox] " & title & ": " & prompt
    InputBoxCI = default
End Function

' Dir - VBAのDir関数（簡易版）
' 注意: VBSにはDir関数がないため、FSOで代替
Function Dir(ByVal pathname)
    If fso_compat.FileExists(pathname) Then
        Dir = fso_compat.GetFileName(pathname)
    ElseIf fso_compat.FolderExists(pathname) Then
        Dir = fso_compat.GetFileName(pathname)
    Else
        Dir = ""
    End If
End Function

' CurDir - 現在のディレクトリを取得
Function CurDir()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    CurDir = shell.CurrentDirectory
End Function

' ChDir - ディレクトリを変更
Sub ChDir(ByVal path)
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    shell.CurrentDirectory = path
End Sub

' Kill - ファイルを削除
Sub Kill(ByVal pathname)
    If fso_compat.FileExists(pathname) Then
        fso_compat.DeleteFile pathname, True
    End If
End Sub

' MkDir - ディレクトリを作成
Sub MkDir(ByVal path)
    If Not fso_compat.FolderExists(path) Then
        fso_compat.CreateFolder path
    End If
End Sub

' RmDir - ディレクトリを削除
Sub RmDir(ByVal path)
    If fso_compat.FolderExists(path) Then
        fso_compat.DeleteFolder path, True
    End If
End Sub

' FileCopy - ファイルをコピー
Sub FileCopy(ByVal source, ByVal destination)
    fso_compat.CopyFile source, destination, True
End Sub

' FileLen - ファイルサイズを取得
Function FileLen(ByVal pathname)
    If fso_compat.FileExists(pathname) Then
        FileLen = fso_compat.GetFile(pathname).Size
    Else
        FileLen = 0
    End If
End Function

' GetAttr - ファイル属性を取得
Function GetAttr(ByVal pathname)
    If fso_compat.FileExists(pathname) Then
        GetAttr = fso_compat.GetFile(pathname).Attributes
    ElseIf fso_compat.FolderExists(pathname) Then
        GetAttr = fso_compat.GetFolder(pathname).Attributes
    Else
        GetAttr = 0
    End If
End Function

' Sgn - 符号を取得
Function Sgn(ByVal number)
    If number > 0 Then
        Sgn = 1
    ElseIf number < 0 Then
        Sgn = -1
    Else
        Sgn = 0
    End If
End Function

' Fix - 整数部分を取得（負の場合は切り上げ）
Function Fix(ByVal number)
    If number >= 0 Then
        Fix = Int(number)
    Else
        Fix = -Int(-number)
    End If
End Function

' Rnd - 乱数（VBSにも存在するが念のため）
' VBSのRndはシードの初期化が必要
Randomize

' ============================================
' 追加の文字列・配列操作関数
' ============================================

' StrComp - 文字列比較（VBSにも存在するが念のため）
' VBSのStrCompは使える

' InStrRev - 後ろから検索（VBSにも存在するが念のため）
' VBSのInStrRevは使える

' StrReverse - 文字列を反転
Function StrReverse(ByVal str)
    Dim i, result
    result = ""
    For i = Len(str) To 1 Step -1
        result = result & Mid(str, i, 1)
    Next
    StrReverse = result
End Function

' Filter - 配列のフィルタリング（VBSにも存在）
' VBSのFilterは使える

' ============================================
' 日付・時刻関数（VBSで使えるもの）
' ============================================
' DateSerial, TimeSerial, DateAdd, DateDiff, DatePart は VBS標準で使える
' Year, Month, Day, Hour, Minute, Second, Weekday は VBS標準で使える
' Now, Date, Time は VBS標準で使える
' Format は FormatDateTime, FormatNumber, FormatCurrency, FormatPercent を使う

' DateValue - 文字列から日付を取得（VBSにも存在）
' TimeValue - 文字列から時刻を取得（VBSにも存在）

' ============================================
' 数学関数（VBSで使えるもの）
' ============================================
' Abs, Sqr, Log, Exp, Sin, Cos, Tan, Atn は VBS標準で使える
' Int, Fix は VBS標準で使える（Fixは上で追加済み）

' Round - 四捨五入（VBSにも存在、銀行家の丸め）
' VBSのRoundは使える

' ============================================
' 型チェック関数（VBSで使えるもの）
' ============================================
' IsArray, IsDate, IsEmpty, IsNull, IsNumeric, IsObject は VBS標準で使える
' VarType, TypeName は VBS標準で使える

' ============================================
' その他の関数
' ============================================

' Switch - 条件分岐（VBSにも存在）
' Choose - 選択（VBSにも存在）

' Asc - 文字コード取得（VBSにも存在）
' AscW - Unicode文字コード取得（VBSにも存在）
' Chr - 文字コードから文字（VBSにも存在）
' ChrW - Unicode文字コードから文字（VBSにも存在）

' Val - 文字列から数値（VBSでは使えない場合がある）
Function Val(ByVal str)
    Dim result, i, c, started, hasDecimal
    result = ""
    started = False
    hasDecimal = False

    str = Trim(str)
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
            started = True
        ElseIf c = "-" And Not started Then
            result = result & c
        ElseIf c = "+" And Not started Then
            ' Skip
        ElseIf c = "." And Not hasDecimal Then
            result = result & c
            hasDecimal = True
            started = True
        ElseIf c = " " And Not started Then
            ' Skip leading spaces
        Else
            Exit For
        End If
    Next

    If result = "" Or result = "-" Or result = "." Then
        Val = 0
    Else
        Val = CDbl(result)
    End If
End Function

' Nz - Null を別の値に変換（Access VBA用、VBSでは使えない）
Function Nz(ByVal value, ByVal valueIfNull)
    If IsNull(value) Or IsEmpty(value) Then
        Nz = valueIfNull
    Else
        Nz = value
    End If
End Function

' IIf - 条件演算子（VBSにも存在）
' VBSのIIfは使える

' ============================================
' ファイルパス操作関数
' ============================================

' App.Path の代替（VBAのApplication.Path）
' GetScriptDir() で代替

' GetTempPath - 一時フォルダのパスを取得
Function GetTempPath()
    Dim shell
    Set shell = CreateObject("WScript.Shell")
    GetTempPath = shell.ExpandEnvironmentStrings("%TEMP%")
    If Right(GetTempPath, 1) <> "\" Then
        GetTempPath = GetTempPath & "\"
    End If
End Function

' ============================================
' VBA定数（VBSでは定義されていないもの）
' ============================================
Const vbObjectError = -2147221504

' ファイル属性定数
Const vbNormal = 0
Const vbReadOnly = 1
Const vbHidden = 2
Const vbSystem = 4
Const vbVolume = 8
Const vbDirectory = 16
Const vbArchive = 32

' 比較モード定数（StrComp用）
Const vbBinaryCompare = 0
Const vbTextCompare = 1

' 日付定数
Const vbSunday = 1
Const vbMonday = 2
Const vbTuesday = 3
Const vbWednesday = 4
Const vbThursday = 5
Const vbFriday = 6
Const vbSaturday = 7

' 第1週の定義
Const vbUseSystem = 0
Const vbFirstJan1 = 1
Const vbFirstFourDays = 2
Const vbFirstFullWeek = 3

' VarType定数（VBSでも使えるが念のため）
Const vbEmpty = 0
Const vbNull = 1
Const vbInteger = 2
Const vbLong = 3
Const vbSingle = 4
Const vbDouble = 5
Const vbCurrency = 6
Const vbDate = 7
Const vbString = 8
Const vbObject = 9
Const vbError = 10
Const vbBoolean = 11
Const vbVariant = 12
Const vbDataObject = 13
Const vbDecimal = 14
Const vbByte = 17
Const vbArray = 8192

' MsgBox定数（VBSでも使えるが念のため）
Const vbOKOnly = 0
Const vbOKCancel = 1
Const vbAbortRetryIgnore = 2
Const vbYesNoCancel = 3
Const vbYesNo = 4
Const vbRetryCancel = 5
Const vbCritical = 16
Const vbQuestion = 32
Const vbExclamation = 48
Const vbInformation = 64
Const vbDefaultButton1 = 0
Const vbDefaultButton2 = 256
Const vbDefaultButton3 = 512
Const vbOK = 1
Const vbCancel = 2
Const vbAbort = 3
Const vbRetry = 4
Const vbIgnore = 5
Const vbYes = 6
Const vbNo = 7

' 改行定数（VBSでも使えるが念のため）
' vbCr, vbLf, vbCrLf, vbNewLine, vbTab は VBS標準で使える
