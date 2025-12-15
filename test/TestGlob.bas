Attribute VB_Name = "TestGlob"
Option Explicit

' ============================================
' SetType / GetType / GetTypeName テスト
' ============================================

Public Sub Test_SetType_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    expect(g.GetTypeName).toBe "Dictionary"
    expect(g.GetType).toBe GlobDataType.dictionary
End Sub

Public Sub Test_SetType_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    expect(g.GetTypeName).toBe "String"
    expect(g.GetType).toBe GlobDataType.StringArray
End Sub

Public Sub Test_SetType_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    expect(g.GetTypeName).toBe "ArrayList"
    expect(g.GetType).toBe GlobDataType.ArrayList
End Sub

' ============================================
' GetCount テスト（初期状態）
' ============================================

Public Sub Test_GetCount_Empty_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    expect(g.GetCount).toBe 0
End Sub

Public Sub Test_GetCount_Empty_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    expect(g.GetCount).toBe 0
End Sub

Public Sub Test_GetCount_Empty_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    expect(g.GetCount).toBe 0
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

    expect(g.GetCount).toBe 3
End Sub

Public Sub Test_AddItem_StringArray()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.StringArray

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"
    g.AddItem "key3", "value3"
    g.AddItem "key4", "value4"

    expect(g.GetCount).toBe 4
End Sub

Public Sub Test_AddItem_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"

    expect(g.GetCount).toBe 2
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

    expect(items.Count).toBe 2
    expect(items("keyA")).toBe "valueA"
    expect(items("keyB")).toBe "valueB"
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

    expect(UBound(items) - LBound(items) + 1).toBe 3
    expect(items(0)).toBe "val1"
    expect(items(1)).toBe "val2"
    expect(items(2)).toBe "val3"
End Sub

Public Sub Test_GetItems_ArrayList()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.ArrayList

    g.AddItem "k1", "v1"
    g.AddItem "k2", "v2"

    Dim items As Object
    Set items = g.GetItems

    expect(items.Count).toBe 2
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
    expect(g.GetCount).toBe 2

    g.Clear

    expect(g.GetCount).toBe 0
    expect(g.GetTypeName).toBe "Dictionary"
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
    expect(g.GetCount).toBe 1

    ' StringArray に変更（データはリセットされる）
    g.SetType = GlobDataType.StringArray

    expect(g.GetTypeName).toBe "String"
    expect(g.GetCount).toBe 0
End Sub

' ============================================
' iGlob ファイル検索テスト
' 注意: iGlob内でMe.Clearが呼ばれ、デフォルトでDictionaryが使われる
' VBSでもFor Eachでイテレート可能
' ============================================

Public Sub Test_iGlob_FindClsFiles()
    Dim g As Glob
    Set g = New Glob

    Dim result As Variant
    Set result = g.iGlob(ThisWorkbook.path & "\..\..\src\*.cls")

    ' Calculator.cls, Glob.cls, WebAPI.cls が存在するはず
    expect(g.GetCount).toBeGreaterThanOrEqual 2
End Sub

Public Sub Test_iGlob_FindBasFiles()
    Dim g As Glob
    Set g = New Glob

    Dim result As Variant
    Set result = g.iGlob(ThisWorkbook.path & "\..\..\test\*.bas")

    ' TestCalculator.bas, TestGlob.bas, TestWebAPI.bas が存在するはず
    expect(g.GetCount).toBeGreaterThanOrEqual 2
End Sub

Public Sub Test_iGlob_WithForEach()
    Dim g As Glob
    Set g = New Glob

    Dim result As Variant
    Set result = g.iGlob(ThisWorkbook.path & "\..\..\src\*.cls")

    Dim count As Long
    Dim item As Variant
    count = 0
    For Each item In result
        count = count + 1
    Next

    expect(count).toBeGreaterThanOrEqual 2
End Sub
