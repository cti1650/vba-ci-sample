VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Glob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DefPath As String
Private Items As Variant
Private FSO As Object

Enum GlobDataType
    None = 0
    StringArray = 1
    ArrayList = 2
    dictionary = 3
    Collection = 4
End Enum

Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    With CreateObject("WScript.Shell")
        .CurrentDirectory = ThisWorkbook.path & "\"
    End With
    Me.Clear
End Sub

Private Sub Class_Terminate()
    Set Items = Nothing
    Set FSO = Nothing
End Sub

Public Sub Clear()

    DefPath = ThisWorkbook.path & "\"
    Select Case Me.GetType
    Case GlobDataType.dictionary
        Me.SetType = GlobDataType.dictionary
    Case GlobDataType.Collection
        Me.SetType = GlobDataType.Collection
    Case GlobDataType.StringArray
        Me.SetType = GlobDataType.StringArray
    Case GlobDataType.ArrayList
        Me.SetType = GlobDataType.ArrayList
    Case Else
        ' VBSではCollectionモックがFor Eachに対応していないため、Dictionaryをデフォルトにする
        Me.SetType = GlobDataType.dictionary
    End Select

End Sub

Public Function GetItems() As Variant
    Select Case Me.GetType
    Case GlobDataType.dictionary, GlobDataType.Collection, GlobDataType.ArrayList
        Set GetItems = Items
    Case GlobDataType.StringArray
        GetItems = Split(Items, "||")
    Case Else
        GetItems = Array()
    End Select
End Function

Public Function GetCount() As Long
    Select Case Me.GetType
    Case GlobDataType.dictionary, GlobDataType.Collection, GlobDataType.ArrayList
        GetCount = Items.count
    Case GlobDataType.StringArray
        If Items = "" Then
            GetCount = 0
        Else
            GetCount = UBound(Split(Items, "||")) + 1
        End If
    Case Else
        GetCount = -1
    End Select
End Function

Public Sub AddItem(ByVal name As String, ByVal v As Variant)
    Select Case Me.GetType
    Case GlobDataType.dictionary
        Items.Add name, v
    Case GlobDataType.Collection
        Items.Add v, name
    Case GlobDataType.ArrayList
        Items.Add v
    Case GlobDataType.StringArray
        If Items <> "" Then Items = Items & "||"
        Items = Items & v
    End Select
End Sub

Public Property Get GetType() As GlobDataType
    Select Case Me.GetTypeName
    Case "Collection"
        GetType = GlobDataType.Collection
    Case "Dictionary"
        GetType = GlobDataType.dictionary
    Case "String"
        GetType = GlobDataType.StringArray
    Case "ArrayList"
        GetType = GlobDataType.ArrayList
    Case Else
        GetType = GlobDataType.None
    End Select
End Property

Public Property Let SetType(ByVal TypeName As GlobDataType)
    Select Case TypeName
    Case GlobDataType.Collection
        Set Items = Nothing
        Set Items = New Collection
    Case GlobDataType.dictionary
        Set Items = Nothing
        Set Items = CreateObject("scripting.dictionary")
    Case GlobDataType.StringArray
        Items = ""
    Case GlobDataType.ArrayList
        Set Items = Nothing
        Set Items = CreateObject("System.Collections.ArrayList")
    Case Else
        Set Items = Nothing
        Set Items = CreateObject("scripting.dictionary")
    End Select
End Property

Public Function GetTypeName() As String
    GetTypeName = TypeName(Items)
End Function


Private Function base(ByRef url As String, Optional ByRef key As String = "") As String
    Dim baseUrl As String
    Dim min As Long
    Dim keystr As String

    If Left$(url, 2) <> "\\" And Left$(url, 1) = "\" Then url = Mid$(url, 2, Len(url) - 1)

    If url <> "" Then
        min = 2000
        If InStr(url, "?") And min > InStr(url, "?") Then min = InStr(url, "?")
        If InStr(url, "*") And min > InStr(url, "*") Then min = InStr(url, "*")
        If InStr(url, "[") And min > InStr(url, "[") Then min = InStr(url, "[")
        If InStr(url, "{") And min > InStr(url, "{") Then min = InStr(url, "{")
        If InStr(url, "]") And min > InStr(url, "]") Then min = InStr(url, "]")
        If InStr(url, "}") And min > InStr(url, "}") Then min = InStr(url, "}")
        If min < 2000 Then
            keystr = Left$(Left$(url, min - 1), InStrRev(Left$(url, min - 1), "\"))
            baseUrl = FSO.GetAbsolutePathName(keystr)
            key = Replace$(url, keystr, "")
        Else
            baseUrl = FSO.GetAbsolutePathName(url)
            key = ""
        End If
        If FSO.FolderExists(baseUrl) = True Then
            url = baseUrl
            base = baseUrl
        Else
            url = ""
            base = ""
        End If
    Else
        url = ""
        key = ""
        base = ""
    End If
End Function

Public Function iGlob(Optional ByVal url As String = "") As Variant

    Dim key As String
    key = ""
    Call base(url, key)
    Me.Clear
    Call subSearch(url, key, 0)
    If IsObject(Me.GetItems) = True Then
        Set iGlob = Me.GetItems
    Else
        iGlob = Me.GetItems
    End If

End Function

Public Function Glob(Optional ByVal url As String = "") As Object

    With New Glob
        .SetType = GlobDataType.dictionary
        Set Glob = .iGlob(url)
    End With

End Function

Public Function GlobFolder(Optional ByVal url As String = "") As Object

    Dim item As Variant
    Dim List As Object
    Set List = CreateObject("scripting.dictionary")
    With New Glob
        .SetType = GlobDataType.Collection
        For Each item In Me.iGlob(url)
            If TypeName(item) = "File" Then
                If List.Exists(item.ParentFolder) = False Then
                    List.Add item.ParentFolder, item.ParentFolder
                End If
            Else
                If List.Exists(item.path) = False Then
                    List.Add item.path, item.path
                End If
            End If
        Next
    End With
    Set GlobFolder = List
    Set List = Nothing

End Function

Private Function subSearch(ByVal baseUrl As String, ByVal key As String, Optional ByVal level As Long = 0) As String

    Dim keyArr As Variant
    Dim folder As Variant
    Dim File As Variant

    keyArr = Split(key, "\")

    If UBound(keyArr) > level Then

        If keyArr(level) = "**" Then
            Call recursive(baseUrl, key, level + 1)
        ElseIf keyArr(level) Like "{*}" Then

            For Each folder In FSO.GetFolder(baseUrl).SubFolders
                If folder.name Like keyArr(level) Then
                    Call subSearch(baseUrl & "\" & folder.name, key, level + 1)
                End If
            Next
            Call subSearch(baseUrl, key, level + 1)

        Else

            For Each folder In FSO.GetFolder(baseUrl).SubFolders
                If folder.name Like keyArr(level) Then
                    Call subSearch(baseUrl & "\" & folder.name, key, level + 1)
                End If
            Next

        End If

    Else

        If keyArr(level) = "" Then

            If FSO.FolderExists(baseUrl) = True Then
                Me.AddItem baseUrl, FSO.GetFolder(baseUrl)
            End If

        Else

            For Each File In FSO.GetFolder(baseUrl).Files
                If File.name Like keyArr(level) Then
                    Me.AddItem File, File
                End If
            Next

        End If

    End If



End Function

Private Function recursive(ByVal baseUrl As String, ByVal key As String, Optional ByVal level As Long = 0) As String

    Dim folder As Variant
    Dim keyArr As Variant
    Dim File As Variant

    keyArr = Split(key, "\")

    If UBound(keyArr) > level Then

        For Each folder In FSO.GetFolder(baseUrl).SubFolders
            If folder.name Like keyArr(level) Then
                Call subSearch(baseUrl & "\" & folder.name, key, level + 1)
            ElseIf "{" & folder.name & "}" Like keyArr(level) Then
                Call subSearch(baseUrl, key, level)
            Else
                Call recursive(baseUrl & "\" & folder.name, key, level)
            End If
        Next

    Else
        For Each folder In FSO.GetFolder(baseUrl).SubFolders
            Call recursive(baseUrl & "\" & folder.name, key, level)
        Next
        For Each File In FSO.GetFolder(baseUrl).Files
            If File.name Like keyArr(level) Then
                Me.AddItem File, File
            End If
        Next
    End If

End Function
