' Expect.vbs - vitest-like assertion class
' Usage: expect(actual).toBe(expected)
'        expect(actual).toEqual(expected)
'        expect(actual).toBeTruthy()
'        expect(actual).toBeFalsy()
'        expect(actual).toBeGreaterThan(value)
'        expect(actual).toBeLessThan(value)
'        expect(actual).toContain(substring)

Class Expectation
    Private actual_
    Private negated_

    Public Sub Init(ByVal val)
        actual_ = val
        negated_ = False
    End Sub

    Public Property Get Not_()
        negated_ = True
        Set Not_ = Me
    End Property

    Private Sub Fail(ByVal assertionType, ByVal expected, ByVal showDiff)
        Dim notStr
        If negated_ Then
            notStr = " not "
        Else
            notStr = " "
        End If

        WScript.Echo ""
        WScript.Echo "AssertionError: expected " & FormatValue(actual_) & notStr & assertionType
        If showDiff Then
            WScript.Echo "  - Expected: " & FormatValue(expected)
            WScript.Echo "  + Received: " & FormatValue(actual_)
        End If
        WScript.Echo ""
        WScript.Quit 1
    End Sub

    Public Sub toBe(ByVal expected)
        Dim passed
        passed = (actual_ = expected)
        If negated_ Then passed = Not passed

        If Not passed Then
            If negated_ Then
                Fail "to be " & FormatValue(expected), expected, False
            Else
                Fail "to be " & FormatValue(expected), expected, True
            End If
        End If
        negated_ = False
    End Sub

    Public Sub toEqual(ByVal expected)
        toBe expected
    End Sub

    Public Sub toBeTruthy()
        Dim passed
        passed = CBool(actual_)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be truthy", True, False
        End If
        negated_ = False
    End Sub

    Public Sub toBeFalsy()
        Dim passed
        passed = Not CBool(actual_)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be falsy", False, False
        End If
        negated_ = False
    End Sub

    Public Sub toBeGreaterThan(ByVal value)
        Dim passed
        passed = (actual_ > value)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be greater than " & FormatValue(value), value, True
        End If
        negated_ = False
    End Sub

    Public Sub toBeLessThan(ByVal value)
        Dim passed
        passed = (actual_ < value)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be less than " & FormatValue(value), value, True
        End If
        negated_ = False
    End Sub

    Public Sub toBeGreaterThanOrEqual(ByVal value)
        Dim passed
        passed = (actual_ >= value)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be >= " & FormatValue(value), value, True
        End If
        negated_ = False
    End Sub

    Public Sub toBeLessThanOrEqual(ByVal value)
        Dim passed
        passed = (actual_ <= value)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be <= " & FormatValue(value), value, True
        End If
        negated_ = False
    End Sub

    Public Sub toContain(ByVal substring)
        Dim passed
        passed = (InStr(1, CStr(actual_), CStr(substring), vbTextCompare) > 0)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to contain " & FormatValue(substring), substring, True
        End If
        negated_ = False
    End Sub

    Public Sub toBeNull()
        Dim passed
        passed = IsNull(actual_)
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be Null", Null, False
        End If
        negated_ = False
    End Sub

    Public Sub toBeEmpty()
        Dim passed
        passed = IsEmpty(actual_) Or (CStr(actual_) = "")
        If negated_ Then passed = Not passed

        If Not passed Then
            Fail "to be empty", "", False
        End If
        negated_ = False
    End Sub

    Private Function FormatValue(ByVal val)
        If IsNull(val) Then
            FormatValue = "Null"
        ElseIf IsEmpty(val) Then
            FormatValue = "Empty"
        ElseIf VarType(val) = vbString Then
            FormatValue = """" & val & """"
        Else
            FormatValue = CStr(val)
        End If
    End Function
End Class

' Factory function for vitest-like syntax
Function expect(ByVal actual)
    Dim e
    Set e = New Expectation
    e.Init actual
    Set expect = e
End Function
