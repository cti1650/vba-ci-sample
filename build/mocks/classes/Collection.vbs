' Collection class mock for VBS
' Note: VBS custom classes do not support For Each
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

    ' Get internal Dictionary (for For Each workaround)
    Public Function GetDict()
        Set GetDict = items_
    End Function
End Class
