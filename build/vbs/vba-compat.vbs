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

' ============================================
' WinAPI モック関数
' VBAのDeclare文で宣言されるAPI関数のVBS用モック
' 実際のAPIは呼べないため、テスト用のダミー実装
' ============================================

' --- kernel32.dll ---

' GetTickCount - システム起動からの経過ミリ秒
Function GetTickCount()
    GetTickCount = CLng(Timer() * 1000) Mod 2147483647
End Function

' GetTickCount64 - 64ビット版（VBSではLongの範囲）
Function GetTickCount64()
    GetTickCount64 = CDbl(Timer() * 1000)
End Function

' Sleep はすでに定義済み（WScript.Sleepを使用）

' GetCurrentProcessId - プロセスID取得
Function GetCurrentProcessId()
    ' WScriptのプロセスIDを返す代替
    Dim shell, exec
    Set shell = CreateObject("WScript.Shell")
    GetCurrentProcessId = shell.Run("cmd /c echo %RANDOM%", 0, True)
    GetCurrentProcessId = Int(Rnd() * 32767) + 1000 ' ダミーのプロセスID
End Function

' GetLastError - 最後のエラーコード
Dim mock_LastError
mock_LastError = 0
Function GetLastError()
    GetLastError = mock_LastError
End Function

Sub SetLastError(ByVal dwErrCode)
    mock_LastError = dwErrCode
End Sub

' GetComputerName - コンピュータ名取得
Function GetComputerNameA(ByRef lpBuffer, ByRef nSize)
    Dim shell, name
    Set shell = CreateObject("WScript.Shell")
    name = shell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    lpBuffer = name
    nSize = Len(name)
    GetComputerNameA = 1 ' 成功
End Function

Function GetComputerNameW(ByRef lpBuffer, ByRef nSize)
    GetComputerNameW = GetComputerNameA(lpBuffer, nSize)
End Function

' GetUserName - ユーザー名取得
Function GetUserNameA(ByRef lpBuffer, ByRef nSize)
    Dim shell, name
    Set shell = CreateObject("WScript.Shell")
    name = shell.ExpandEnvironmentStrings("%USERNAME%")
    lpBuffer = name
    nSize = Len(name)
    GetUserNameA = 1 ' 成功
End Function

Function GetUserNameW(ByRef lpBuffer, ByRef nSize)
    GetUserNameW = GetUserNameA(lpBuffer, nSize)
End Function

' GetTempPathA/W - 一時フォルダパス取得
Function GetTempPathA(ByVal nBufferLength, ByRef lpBuffer)
    Dim path
    path = GetTempPath()
    lpBuffer = path
    GetTempPathA = Len(path)
End Function

Function GetTempPathW(ByVal nBufferLength, ByRef lpBuffer)
    GetTempPathW = GetTempPathA(nBufferLength, lpBuffer)
End Function

' GetSystemDirectory - システムディレクトリ取得
Function GetSystemDirectoryA(ByRef lpBuffer, ByVal uSize)
    Dim shell, path
    Set shell = CreateObject("WScript.Shell")
    path = shell.ExpandEnvironmentStrings("%SystemRoot%\System32")
    lpBuffer = path
    GetSystemDirectoryA = Len(path)
End Function

Function GetSystemDirectoryW(ByRef lpBuffer, ByVal uSize)
    GetSystemDirectoryW = GetSystemDirectoryA(lpBuffer, uSize)
End Function

' GetWindowsDirectory - Windowsディレクトリ取得
Function GetWindowsDirectoryA(ByRef lpBuffer, ByVal uSize)
    Dim shell, path
    Set shell = CreateObject("WScript.Shell")
    path = shell.ExpandEnvironmentStrings("%SystemRoot%")
    lpBuffer = path
    GetWindowsDirectoryA = Len(path)
End Function

Function GetWindowsDirectoryW(ByRef lpBuffer, ByVal uSize)
    GetWindowsDirectoryW = GetWindowsDirectoryA(lpBuffer, uSize)
End Function

' QueryPerformanceCounter - 高精度タイマー
Function QueryPerformanceCounter(ByRef lpPerformanceCount)
    lpPerformanceCount = CDbl(Timer() * 1000000)
    QueryPerformanceCounter = 1 ' 成功
End Function

' QueryPerformanceFrequency - タイマー周波数
Function QueryPerformanceFrequency(ByRef lpFrequency)
    lpFrequency = 1000000 ' 1MHz (ダミー)
    QueryPerformanceFrequency = 1 ' 成功
End Function

' CopyMemory/RtlMoveMemory - メモリコピー（モック：何もしない）
Sub CopyMemory(ByRef Destination, ByRef Source, ByVal Length)
    ' VBSではメモリ操作不可、警告のみ
    DebugPrint "[MOCK WARNING] CopyMemory called - operation not supported in VBS"
End Sub

Sub RtlMoveMemory(ByRef Destination, ByRef Source, ByVal Length)
    CopyMemory Destination, Source, Length
End Sub

' --- user32.dll ---

' MessageBox - メッセージボックス（CI用モック）
Function MessageBoxA(ByVal hWnd, ByVal lpText, ByVal lpCaption, ByVal uType)
    DebugPrint "[MessageBox] " & lpCaption & ": " & lpText
    MessageBoxA = 1 ' IDOK
End Function

Function MessageBoxW(ByVal hWnd, ByVal lpText, ByVal lpCaption, ByVal uType)
    MessageBoxW = MessageBoxA(hWnd, lpText, lpCaption, uType)
End Function

' GetActiveWindow - アクティブウィンドウハンドル
Function GetActiveWindow()
    GetActiveWindow = 0 ' ダミーハンドル
End Function

' GetForegroundWindow - フォアグラウンドウィンドウハンドル
Function GetForegroundWindow()
    GetForegroundWindow = 0 ' ダミーハンドル
End Function

' FindWindow - ウィンドウ検索
Function FindWindowA(ByVal lpClassName, ByVal lpWindowName)
    FindWindowA = 0 ' 見つからない
End Function

Function FindWindowW(ByVal lpClassName, ByVal lpWindowName)
    FindWindowW = 0
End Function

' GetWindowText - ウィンドウテキスト取得
Function GetWindowTextA(ByVal hWnd, ByRef lpString, ByVal nMaxCount)
    lpString = ""
    GetWindowTextA = 0
End Function

Function GetWindowTextW(ByVal hWnd, ByRef lpString, ByVal nMaxCount)
    GetWindowTextW = GetWindowTextA(hWnd, lpString, nMaxCount)
End Function

' SetWindowText - ウィンドウテキスト設定
Function SetWindowTextA(ByVal hWnd, ByVal lpString)
    SetWindowTextA = 0 ' 失敗
End Function

Function SetWindowTextW(ByVal hWnd, ByVal lpString)
    SetWindowTextW = 0
End Function

' GetCursorPos - カーソル位置取得
Function GetCursorPos(ByRef lpPoint)
    ' lpPointはオブジェクトまたは配列を想定
    ' VBSでは構造体がないのでダミー値
    GetCursorPos = 0
End Function

' SetCursorPos - カーソル位置設定
Function SetCursorPos(ByVal X, ByVal Y)
    SetCursorPos = 0
End Function

' GetAsyncKeyState - キー状態取得
Function GetAsyncKeyState(ByVal vKey)
    GetAsyncKeyState = 0 ' キーは押されていない
End Function

' GetKeyState - キー状態取得
Function GetKeyState(ByVal nVirtKey)
    GetKeyState = 0
End Function

' SendMessage - メッセージ送信（モック）
Function SendMessageA(ByVal hWnd, ByVal Msg, ByVal wParam, ByVal lParam)
    SendMessageA = 0
End Function

Function SendMessageW(ByVal hWnd, ByVal Msg, ByVal wParam, ByVal lParam)
    SendMessageW = 0
End Function

' PostMessage - メッセージ投稿（モック）
Function PostMessageA(ByVal hWnd, ByVal Msg, ByVal wParam, ByVal lParam)
    PostMessageA = 0
End Function

Function PostMessageW(ByVal hWnd, ByVal Msg, ByVal wParam, ByVal lParam)
    PostMessageW = 0
End Function

' --- shell32.dll ---

' SHGetFolderPath - 特殊フォルダパス取得
Function SHGetFolderPathA(ByVal hwndOwner, ByVal nFolder, ByVal hToken, ByVal dwFlags, ByRef pszPath)
    Dim shell
    Set shell = CreateObject("WScript.Shell")

    Select Case nFolder
        Case 0 ' CSIDL_DESKTOP
            pszPath = shell.SpecialFolders("Desktop")
        Case 5 ' CSIDL_PERSONAL (My Documents)
            pszPath = shell.SpecialFolders("MyDocuments")
        Case 26 ' CSIDL_APPDATA
            pszPath = shell.ExpandEnvironmentStrings("%APPDATA%")
        Case 28 ' CSIDL_LOCAL_APPDATA
            pszPath = shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
        Case 35 ' CSIDL_COMMON_DOCUMENTS
            pszPath = shell.ExpandEnvironmentStrings("%PUBLIC%\Documents")
        Case 36 ' CSIDL_PROGRAM_FILES
            pszPath = shell.ExpandEnvironmentStrings("%ProgramFiles%")
        Case 37 ' CSIDL_WINDOWS
            pszPath = shell.ExpandEnvironmentStrings("%SystemRoot%")
        Case Else
            pszPath = ""
    End Select

    If pszPath <> "" Then
        SHGetFolderPathA = 0 ' S_OK
    Else
        SHGetFolderPathA = 1 ' エラー
    End If
End Function

Function SHGetFolderPathW(ByVal hwndOwner, ByVal nFolder, ByVal hToken, ByVal dwFlags, ByRef pszPath)
    SHGetFolderPathW = SHGetFolderPathA(hwndOwner, nFolder, hToken, dwFlags, pszPath)
End Function

' --- ole32.dll / oleaut32.dll ---

' CoCreateGuid - GUID生成
Function CoCreateGuid(ByRef pguid)
    ' VBSでGUID生成
    Dim typeLib
    Set typeLib = CreateObject("Scriptlet.TypeLib")
    pguid = Mid(typeLib.GUID, 2, 36)
    CoCreateGuid = 0 ' S_OK
End Function

' --- advapi32.dll ---

' GetUserName は上で定義済み

' RegOpenKeyEx - レジストリキーを開く（モック）
Function RegOpenKeyExA(ByVal hKey, ByVal lpSubKey, ByVal ulOptions, ByVal samDesired, ByRef phkResult)
    phkResult = 0
    RegOpenKeyExA = 2 ' ERROR_FILE_NOT_FOUND
End Function

' RegQueryValueEx - レジストリ値を取得（モック）
Function RegQueryValueExA(ByVal hKey, ByVal lpValueName, ByVal lpReserved, ByRef lpType, ByRef lpData, ByRef lpcbData)
    RegQueryValueExA = 2 ' ERROR_FILE_NOT_FOUND
End Function

' RegCloseKey - レジストリキーを閉じる
Function RegCloseKey(ByVal hKey)
    RegCloseKey = 0 ' ERROR_SUCCESS
End Function

' ============================================
' CreateObject モック/ラッパー
' テスト用にCreateObjectをインターセプトして
' モックオブジェクトを返す機能
' ============================================

' モックオブジェクト管理用Dictionary
Dim mock_Objects
Set mock_Objects = CreateObject("Scripting.Dictionary")

' モック登録関数
Sub RegisterMockObject(ByVal progId, ByVal mockObj)
    mock_Objects(LCase(progId)) = mockObj
End Sub

' モック解除関数
Sub UnregisterMockObject(ByVal progId)
    If mock_Objects.Exists(LCase(progId)) Then
        mock_Objects.Remove LCase(progId)
    End If
End Sub

' 全モッククリア
Sub ClearAllMocks()
    mock_Objects.RemoveAll
End Sub

' CreateObjectのラッパー（モック対応）
Function CreateObjectMock(ByVal progId)
    Dim lowerProgId
    lowerProgId = LCase(progId)

    ' モックが登録されていればそれを返す
    If mock_Objects.Exists(lowerProgId) Then
        Set CreateObjectMock = mock_Objects(lowerProgId)
        Exit Function
    End If

    ' モックがなければ実際のオブジェクトを生成
    Set CreateObjectMock = CreateObject(progId)
End Function

' ============================================
' Excel Application モック
' ============================================
Class MockExcelApplication
    Private worksheets_
    Private workbooks_
    Private visible_
    Private displayAlerts_
    Private screenUpdating_
    Private calculation_

    Private Sub Class_Initialize()
        Set worksheets_ = CreateObject("Scripting.Dictionary")
        Set workbooks_ = CreateObject("Scripting.Dictionary")
        visible_ = False
        displayAlerts_ = True
        screenUpdating_ = True
        calculation_ = -4105 ' xlCalculationAutomatic
    End Sub

    Public Property Get Visible()
        Visible = visible_
    End Property

    Public Property Let Visible(ByVal value)
        visible_ = value
    End Property

    Public Property Get DisplayAlerts()
        DisplayAlerts = displayAlerts_
    End Property

    Public Property Let DisplayAlerts(ByVal value)
        displayAlerts_ = value
    End Property

    Public Property Get ScreenUpdating()
        ScreenUpdating = screenUpdating_
    End Property

    Public Property Let ScreenUpdating(ByVal value)
        screenUpdating_ = value
    End Property

    Public Property Get Calculation()
        Calculation = calculation_
    End Property

    Public Property Let Calculation(ByVal value)
        calculation_ = value
    End Property

    Public Property Get Workbooks()
        Set Workbooks = workbooks_
    End Property

    Public Function Quit()
        DebugPrint "[MockExcel] Application.Quit called"
    End Function

    Public Function Run(ByVal macroName)
        DebugPrint "[MockExcel] Application.Run: " & macroName
        Run = Empty
    End Function

    Public Property Get Version()
        Version = "16.0" ' Excel 2016+
    End Property

    Public Property Get Name()
        Name = "Microsoft Excel"
    End Property
End Class

' Excelモック生成関数
Function CreateMockExcelApplication()
    Set CreateMockExcelApplication = New MockExcelApplication
End Function

' ============================================
' Word Application モック
' ============================================
Class MockWordApplication
    Private documents_
    Private visible_
    Private displayAlerts_

    Private Sub Class_Initialize()
        Set documents_ = CreateObject("Scripting.Dictionary")
        visible_ = False
        displayAlerts_ = 0
    End Sub

    Public Property Get Visible()
        Visible = visible_
    End Property

    Public Property Let Visible(ByVal value)
        visible_ = value
    End Property

    Public Property Get DisplayAlerts()
        DisplayAlerts = displayAlerts_
    End Property

    Public Property Let DisplayAlerts(ByVal value)
        displayAlerts_ = value
    End Property

    Public Property Get Documents()
        Set Documents = documents_
    End Property

    Public Function Quit()
        DebugPrint "[MockWord] Application.Quit called"
    End Function

    Public Property Get Version()
        Version = "16.0"
    End Property

    Public Property Get Name()
        Name = "Microsoft Word"
    End Property
End Class

Function CreateMockWordApplication()
    Set CreateMockWordApplication = New MockWordApplication
End Function

' ============================================
' ADODB.Connection モック
' ============================================
Class MockADODBConnection
    Private connectionString_
    Private state_
    Private errors_

    Private Sub Class_Initialize()
        connectionString_ = ""
        state_ = 0 ' adStateClosed
        Set errors_ = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get ConnectionString()
        ConnectionString = connectionString_
    End Property

    Public Property Let ConnectionString(ByVal value)
        connectionString_ = value
    End Property

    Public Property Get State()
        State = state_
    End Property

    Public Property Get Errors()
        Set Errors = errors_
    End Property

    Public Sub Open(ByVal connStr)
        If connStr <> "" Then connectionString_ = connStr
        state_ = 1 ' adStateOpen
        DebugPrint "[MockADODB] Connection.Open: " & connectionString_
    End Sub

    Public Sub Close()
        state_ = 0 ' adStateClosed
        DebugPrint "[MockADODB] Connection.Close"
    End Sub

    Public Function Execute(ByVal commandText)
        DebugPrint "[MockADODB] Connection.Execute: " & commandText
        Set Execute = New MockADODBRecordset
    End Function

    Public Function BeginTrans()
        DebugPrint "[MockADODB] BeginTrans"
        BeginTrans = 1
    End Function

    Public Sub CommitTrans()
        DebugPrint "[MockADODB] CommitTrans"
    End Sub

    Public Sub RollbackTrans()
        DebugPrint "[MockADODB] RollbackTrans"
    End Sub
End Class

' ============================================
' ADODB.Recordset モック
' ============================================
Class MockADODBRecordset
    Private fields_
    Private data_
    Private currentRow_
    Private eof_
    Private bof_

    Private Sub Class_Initialize()
        Set fields_ = CreateObject("Scripting.Dictionary")
        Set data_ = CreateObject("Scripting.Dictionary")
        currentRow_ = -1
        eof_ = True
        bof_ = True
    End Sub

    Public Property Get EOF()
        EOF = eof_
    End Property

    Public Property Get BOF()
        BOF = bof_
    End Property

    Public Property Get Fields()
        Set Fields = fields_
    End Property

    Public Property Get RecordCount()
        RecordCount = data_.Count
    End Property

    Public Sub Open(ByVal source, ByVal conn)
        DebugPrint "[MockADODB] Recordset.Open: " & source
        eof_ = True
        bof_ = True
    End Sub

    Public Sub Close()
        DebugPrint "[MockADODB] Recordset.Close"
    End Sub

    Public Sub MoveFirst()
        If data_.Count > 0 Then
            currentRow_ = 0
            eof_ = False
            bof_ = False
        End If
    End Sub

    Public Sub MoveNext()
        currentRow_ = currentRow_ + 1
        If currentRow_ >= data_.Count Then
            eof_ = True
        End If
    End Sub

    Public Sub MoveLast()
        If data_.Count > 0 Then
            currentRow_ = data_.Count - 1
            eof_ = False
            bof_ = False
        End If
    End Sub

    Public Sub AddNew()
        DebugPrint "[MockADODB] Recordset.AddNew"
    End Sub

    Public Sub Update()
        DebugPrint "[MockADODB] Recordset.Update"
    End Sub

    Public Sub Delete()
        DebugPrint "[MockADODB] Recordset.Delete"
    End Sub
End Class

Function CreateMockADODBConnection()
    Set CreateMockADODBConnection = New MockADODBConnection
End Function

Function CreateMockADODBRecordset()
    Set CreateMockADODBRecordset = New MockADODBRecordset
End Function

' ============================================
' XMLHTTP / ServerXMLHTTP モック
' ============================================
Class MockXMLHTTP
    Private status_
    Private statusText_
    Private responseText_
    Private responseXML_
    Private readyState_
    Private requestMethod_
    Private requestUrl_
    Private requestHeaders_

    Private Sub Class_Initialize()
        status_ = 200
        statusText_ = "OK"
        responseText_ = ""
        Set responseXML_ = Nothing
        readyState_ = 0
        requestMethod_ = ""
        requestUrl_ = ""
        Set requestHeaders_ = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get Status()
        Status = status_
    End Property

    Public Property Get statusText()
        statusText = statusText_
    End Property

    Public Property Get responseText()
        responseText = responseText_
    End Property

    Public Property Get responseXML()
        Set responseXML = responseXML_
    End Property

    Public Property Get readyState()
        readyState = readyState_
    End Property

    Public Sub Open(ByVal method, ByVal url, ByVal async)
        requestMethod_ = method
        requestUrl_ = url
        readyState_ = 1
        DebugPrint "[MockXMLHTTP] Open: " & method & " " & url
    End Sub

    Public Sub setRequestHeader(ByVal header, ByVal value)
        requestHeaders_(header) = value
    End Sub

    Public Sub send(ByVal body)
        readyState_ = 4
        DebugPrint "[MockXMLHTTP] Send: " & requestUrl_
        If body <> "" Then
            DebugPrint "[MockXMLHTTP] Body: " & Left(body, 100)
        End If
        ' デフォルトは空のレスポンス
        responseText_ = "{}"
    End Sub

    Public Function getResponseHeader(ByVal header)
        getResponseHeader = ""
    End Function

    Public Function getAllResponseHeaders()
        getAllResponseHeaders = ""
    End Function

    ' テスト用：レスポンスを設定
    Public Sub SetMockResponse(ByVal statusCode, ByVal text)
        status_ = statusCode
        responseText_ = text
        If statusCode >= 200 And statusCode < 300 Then
            statusText_ = "OK"
        Else
            statusText_ = "Error"
        End If
    End Sub
End Class

Function CreateMockXMLHTTP()
    Set CreateMockXMLHTTP = New MockXMLHTTP
End Function

' ============================================
' RegExp 追加メソッド（VBSのRegExpを拡張）
' ============================================
' VBSのRegExpは標準で使えるので追加定義不要

' ============================================
' WinAPI定数
' ============================================

' CSIDL定数（特殊フォルダ）
Const CSIDL_DESKTOP = 0
Const CSIDL_INTERNET = 1
Const CSIDL_PROGRAMS = 2
Const CSIDL_CONTROLS = 3
Const CSIDL_PRINTERS = 4
Const CSIDL_PERSONAL = 5
Const CSIDL_FAVORITES = 6
Const CSIDL_STARTUP = 7
Const CSIDL_RECENT = 8
Const CSIDL_SENDTO = 9
Const CSIDL_BITBUCKET = 10
Const CSIDL_STARTMENU = 11
Const CSIDL_DESKTOPDIRECTORY = 16
Const CSIDL_DRIVES = 17
Const CSIDL_NETWORK = 18
Const CSIDL_NETHOOD = 19
Const CSIDL_FONTS = 20
Const CSIDL_TEMPLATES = 21
Const CSIDL_COMMON_STARTMENU = 22
Const CSIDL_COMMON_PROGRAMS = 23
Const CSIDL_COMMON_STARTUP = 24
Const CSIDL_COMMON_DESKTOPDIRECTORY = 25
Const CSIDL_APPDATA = 26
Const CSIDL_PRINTHOOD = 27
Const CSIDL_LOCAL_APPDATA = 28
Const CSIDL_COMMON_FAVORITES = 31
Const CSIDL_INTERNET_CACHE = 32
Const CSIDL_COOKIES = 33
Const CSIDL_HISTORY = 34
Const CSIDL_COMMON_APPDATA = 35
Const CSIDL_WINDOWS = 36
Const CSIDL_SYSTEM = 37
Const CSIDL_PROGRAM_FILES = 38
Const CSIDL_MYPICTURES = 39
Const CSIDL_PROFILE = 40
Const CSIDL_PROGRAM_FILES_COMMON = 43
Const CSIDL_COMMON_TEMPLATES = 45
Const CSIDL_COMMON_DOCUMENTS = 46
Const CSIDL_COMMON_ADMINTOOLS = 47
Const CSIDL_ADMINTOOLS = 48
Const CSIDL_COMMON_MUSIC = 53
Const CSIDL_COMMON_PICTURES = 54
Const CSIDL_COMMON_VIDEO = 55
Const CSIDL_CDBURN_AREA = 59

' 仮想キーコード
Const VK_LBUTTON = &H1
Const VK_RBUTTON = &H2
Const VK_CANCEL = &H3
Const VK_MBUTTON = &H4
Const VK_BACK = &H8
Const VK_TAB = &H9
Const VK_CLEAR = &HC
Const VK_RETURN = &HD
Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12
Const VK_PAUSE = &H13
Const VK_CAPITAL = &H14
Const VK_ESCAPE = &H1B
Const VK_SPACE = &H20
Const VK_PRIOR = &H21
Const VK_NEXT = &H22
Const VK_END = &H23
Const VK_HOME = &H24
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28
Const VK_SELECT = &H29
Const VK_PRINT = &H2A
Const VK_EXECUTE = &H2B
Const VK_SNAPSHOT = &H2C
Const VK_INSERT = &H2D
Const VK_DELETE = &H2E
Const VK_HELP = &H2F
Const VK_F1 = &H70
Const VK_F2 = &H71
Const VK_F3 = &H72
Const VK_F4 = &H73
Const VK_F5 = &H74
Const VK_F6 = &H75
Const VK_F7 = &H76
Const VK_F8 = &H77
Const VK_F9 = &H78
Const VK_F10 = &H79
Const VK_F11 = &H7A
Const VK_F12 = &H7B

' メッセージ定数
Const WM_NULL = &H0
Const WM_CREATE = &H1
Const WM_DESTROY = &H2
Const WM_MOVE = &H3
Const WM_SIZE = &H5
Const WM_ACTIVATE = &H6
Const WM_SETFOCUS = &H7
Const WM_KILLFOCUS = &H8
Const WM_ENABLE = &HA
Const WM_SETTEXT = &HC
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const WM_PAINT = &HF
Const WM_CLOSE = &H10
Const WM_QUIT = &H12
Const WM_SHOWWINDOW = &H18
Const WM_SETFONT = &H30
Const WM_GETFONT = &H31
Const WM_COMMAND = &H111
Const WM_USER = &H400

' Excel定数
Const xlCalculationAutomatic = -4105
Const xlCalculationManual = -4135
Const xlCalculationSemiautomatic = 2
