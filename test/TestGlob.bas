Attribute VB_Name = "TestGlob"
Option Explicit

' ============================================
' SetType / GetType / GetTypeName テスト
' ============================================

Public Sub Test_SetType_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    If g.GetTypeName <> "Dictionary" Then
        Utils.Fail 2001, "GetTypeName expected 'Dictionary'"
    End If

    If g.GetType <> GlobDataType.dictionary Then
        Utils.Fail 2002, "GetType expected GlobDataType.dictionary"
    End If
End Sub

Public Sub Test_SetType_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    If g.GetTypeName <> "String" Then
        Utils.Fail 2005, "GetTypeName expected 'String'"
    End If

    If g.GetType <> GlobDataType.StringArray Then
        Utils.Fail 2006, "GetType expected GlobDataType.StringArray"
    End If
End Sub

Public Sub Test_SetType_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    If g.GetTypeName <> "ArrayList" Then
        Utils.Fail 2007, "GetTypeName expected 'ArrayList'"
    End If

    If g.GetType <> GlobDataType.ArrayList Then
        Utils.Fail 2008, "GetType expected GlobDataType.ArrayList"
    End If
End Sub

' ============================================
' GetCount テスト（初期状態）
' ============================================

Public Sub Test_GetCount_Empty_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    If g.GetCount <> 0 Then
        Utils.Fail 2010, "GetCount expected 0 for empty Dictionary"
    End If
End Sub

Public Sub Test_GetCount_Empty_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    If g.GetCount <> 0 Then
        Utils.Fail 2012, "GetCount expected 0 for empty StringArray"
    End If
End Sub

Public Sub Test_GetCount_Empty_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    If g.GetCount <> 0 Then
        Utils.Fail 2013, "GetCount expected 0 for empty ArrayList"
    End If
End Sub

' ============================================
' AddItem / GetCount テスト
' ============================================

Public Sub Test_AddItem_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"
    g.AddItem "key3", "value3"

    If g.GetCount <> 3 Then
        Utils.Fail 2020, "GetCount expected 3 after adding 3 items to Dictionary"
    End If
End Sub

Public Sub Test_AddItem_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"
    g.AddItem "key3", "value3"
    g.AddItem "key4", "value4"

    If g.GetCount <> 4 Then
        Utils.Fail 2022, "GetCount expected 4 after adding 4 items to StringArray"
    End If
End Sub

Public Sub Test_AddItem_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"

    If g.GetCount <> 2 Then
        Utils.Fail 2023, "GetCount expected 2 after adding 2 items to ArrayList"
    End If
End Sub

' ============================================
' GetItems テスト
' ============================================

Public Sub Test_GetItems_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    g.AddItem "keyA", "valueA"
    g.AddItem "keyB", "valueB"

    Dim items As Object
    Set items = g.GetItems

    If items.Count <> 2 Then
        Utils.Fail 2030, "GetItems.Count expected 2 for Dictionary"
    End If

    If items("keyA") <> "valueA" Then
        Utils.Fail 2031, "GetItems('keyA') expected 'valueA'"
    End If
End Sub

Public Sub Test_GetItems_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    g.AddItem "k1", "val1"
    g.AddItem "k2", "val2"
    g.AddItem "k3", "val3"

    Dim items As Variant
    items = g.GetItems

    If UBound(items) - LBound(items) + 1 <> 3 Then
        Utils.Fail 2033, "GetItems array length expected 3 for StringArray"
    End If

    If items(0) <> "val1" Then
        Utils.Fail 2034, "GetItems(0) expected 'val1'"
    End If
End Sub

Public Sub Test_GetItems_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    g.AddItem "k1", "v1"
    g.AddItem "k2", "v2"

    Dim items As Object
    Set items = g.GetItems

    If items.Count <> 2 Then
        Utils.Fail 2035, "GetItems.Count expected 2 for ArrayList"
    End If
End Sub

' ============================================
' Clear テスト
' ============================================

Public Sub Test_Clear_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"

    If g.GetCount <> 2 Then
        Utils.Fail 2040, "GetCount expected 2 before Clear"
    End If

    g.Clear

    If g.GetCount <> 0 Then
        Utils.Fail 2041, "GetCount expected 0 after Clear"
    End If

    ' Clear後も同じ型を維持
    If g.GetTypeName <> "Dictionary" Then
        Utils.Fail 2042, "GetTypeName expected 'Dictionary' after Clear"
    End If
End Sub

' ============================================
' 型変更テスト
' ============================================

Public Sub Test_ChangeType()
    Dim g As Glob
    Set g = New Glob

    ' Dictionary に設定
    g.SetType = GlobDataType.dictionary
    g.AddItem "key1", "value1"

    If g.GetCount <> 1 Then
        Utils.Fail 2050, "GetCount expected 1 for Dictionary"
    End If

    ' StringArray に変更（データはリセットされる）
    g.SetType = GlobDataType.StringArray

    If g.GetTypeName <> "String" Then
        Utils.Fail 2051, "GetTypeName expected 'String' after type change"
    End If

    If g.GetCount <> 0 Then
        Utils.Fail 2052, "GetCount expected 0 after type change (data should be reset)"
    End If
End Sub

' ============================================
' iGlob ファイル検索テスト
' 注意: iGlob内でMe.Clearが呼ばれ、デフォルトでDictionaryが使われる
' VBSでもFor Eachでイテレート可能
' ============================================

Public Sub Test_iGlob_FindClsFiles()
    ' src/*.cls ファイルを検索（CI環境: build/vbs/ から実行）
    Dim g As Glob
    Set g = New Glob

    ' iGlob内でClearが呼ばれSetTypeがCollectionになる
    ' GetScriptDir() は build/vbs を返すので、../../src/*.cls でsrcを参照
    Dim result As Variant
    Set result = g.iGlob(ThisWorkbook.path & "\..\..\src\*.cls")

    ' Calculator.cls と Glob.cls が存在するはず
    If g.GetCount < 2 Then
        Utils.Fail 2060, "iGlob should find at least 2 .cls files in src/"
    End If
End Sub

Public Sub Test_iGlob_FindBasFiles()
    ' test/*.bas ファイルを検索
    Dim g As Glob
    Set g = New Glob

    Dim result As Variant
    Set result = g.iGlob(ThisWorkbook.path & "\..\..\test\*.bas")

    ' TestCalculator.bas と TestGlob.bas が存在するはず
    If g.GetCount = 2 Then
        Utils.Fail 2061, "iGlob should find at least 2 .bas files in test/"
    End If
End Sub

Public Sub Test_iGlob_WithForEach()
    ' For Each で結果をイテレートできることを確認
    Dim g As Glob
    Set g = New Glob

    Dim result As Variant
    Set result = g.iGlob(ThisWorkbook.path & "\..\..\src\*.cls")

    ' デフォルトでDictionaryが使われるため、For Eachでイテレート可能
    Dim count As Long
    Dim item As Variant
    count = 0
    For Each item In result
        count = count + 1
    Next

    If count < 2 Then
        Utils.Fail 2062, "For Each should iterate at least 2 items"
    End If
End Sub
