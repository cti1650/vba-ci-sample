Attribute VB_Name = "TestGlob"
Option Explicit

Public Sub Test_GetCount_Empty()
    Dim g As Glob
    Set g = New Glob

    If g.GetCount <> 0 Then
        Utils.Fail 2001, "GetCount expected 0 for empty Glob"
    End If
End Sub

Public Sub Test_SetType_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    If g.GetTypeName <> "Dictionary" Then
        Utils.Fail 2002, "GetTypeName expected 'Dictionary'"
    End If
End Sub

Public Sub Test_SetType_Collection()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.Collection

    If g.GetTypeName <> "Collection" Then
        Utils.Fail 2003, "GetTypeName expected 'Collection'"
    End If
End Sub

Public Sub Test_AddItem_Dictionary()
    Dim g As Glob
    Set g = New Glob
    g.SetType = GlobDataType.dictionary

    g.AddItem "key1", "value1"
    g.AddItem "key2", "value2"

    If g.GetCount <> 2 Then
        Utils.Fail 2004, "GetCount expected 2 after adding 2 items"
    End If
End Sub

Public Sub Test_Glob_CurrentFolder()
    Dim g As Glob
    Dim result As Object
    Set g = New Glob

    ' VBSでは Glob() メソッド内で New Glob ができないため、
    ' iGlob を直接使用（SetType を先に設定）
    ' パスは vbs フォルダからの相対パス（..で親ディレクトリへ）
    g.SetType = GlobDataType.dictionary
    Set result = g.iGlob(ThisWorkbook.path & "\..\vba\src\*.cls")

    If result.Count < 1 Then
        Utils.Fail 2005, "Glob should find at least 1 .cls file"
    End If
End Sub
