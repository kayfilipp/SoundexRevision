Attribute VB_Name = "Formula"

Function rollingSim(baseStr As String, searchStr As String) As Double
    Dim rollLength As Long
    Dim bestSim As Double
    Dim crossSection As String
    Dim crossSim As Double
    
    'a crucial condition here is that the base string must be larger than the search string. 
	'if we do not satisfy it, we revert back to the regular similarity function.
    If Len(searchStr) >= Len(baseStr) Then
        rollingSim = Utils.Similarity(searchStr, baseStr)
        Exit Function
    End If
    
    'how long each mid-section of the base string is.
    rollLength = Len(searchStr)
    'the closest levenshtein distance of all possible mid-sections
    bestSim = 0
    
    For i = 1 To Len(baseStr) - rollLength + 1 '+1 to offset the possibility of two strings being the same.
        crossSection = Mid(baseStr, i, rollLength)
        crossSim = Utils.Similarity(crossSection, searchStr)
        If crossSim > bestSim Then
            bestSim = crossSim
        End If
    Next
    
    rollingSim = bestSim
End Function

Public Function Similarity(ByVal String1 As String, _
    ByVal String2 As String, _
    Optional ByRef RetMatch As String, _
    Optional min_match = 1) As Single
    Dim b1() As Byte, b2() As Byte
    Dim lngLen1 As Long, lngLen2 As Long
    Dim lngResult As Long

    If UCase(String1) = UCase(String2) Then
        Similarity = 1
    Else:
        lngLen1 = Len(String1)
        lngLen2 = Len(String2)
        If (lngLen1 = 0) Or (lngLen2 = 0) Then
            Similarity = 0
        Else:
            b1() = StrConv(UCase(String1), vbFromUnicode)
            b2() = StrConv(UCase(String2), vbFromUnicode)
            lngResult = Similarity_sub(0, lngLen1 - 1, _
            0, lngLen2 - 1, _
            b1, b2, _
            String1, _
            RetMatch, _
            min_match)
            Erase b1
            Erase b2
            If lngLen1 >= lngLen2 Then
                Similarity = lngResult / lngLen1
            Else
                Similarity = lngResult / lngLen2
            End If
        End If
    End If

End Function

Private Function Similarity_sub(ByVal start1 As Long, ByVal end1 As Long, _
                                ByVal start2 As Long, ByVal end2 As Long, _
                                ByRef b1() As Byte, ByRef b2() As Byte, _
                                ByVal FirstString As String, _
                                ByRef RetMatch As String, _
                                ByVal min_match As Long, _
                                Optional recur_level As Integer = 0) As Long
'* CALLED BY: Similarity *(RECURSIVE)

    Dim lngCurr1 As Long, lngCurr2 As Long
    Dim lngMatchAt1 As Long, lngMatchAt2 As Long
    Dim i As Long
    Dim lngLongestMatch As Long, lngLocalLongestMatch As Long
    Dim strRetMatch1 As String, strRetMatch2 As String

    If (start1 > end1) Or (start1 < 0) Or (end1 - start1 + 1 < min_match) _
    Or (start2 > end2) Or (start2 < 0) Or (end2 - start2 + 1 < min_match) Then
        Exit Function '(exit if start/end is out of string, or length is too short)
    End If

    For lngCurr1 = start1 To end1
        For lngCurr2 = start2 To end2
            i = 0
            Do Until b1(lngCurr1 + i) <> b2(lngCurr2 + i)
                i = i + 1
                If i > lngLongestMatch Then
                    lngMatchAt1 = lngCurr1
                    lngMatchAt2 = lngCurr2
                    lngLongestMatch = i
                End If
                If (lngCurr1 + i) > end1 Or (lngCurr2 + i) > end2 Then Exit Do
            Loop
        Next lngCurr2
    Next lngCurr1

    If lngLongestMatch < min_match Then Exit Function

    lngLocalLongestMatch = lngLongestMatch
    RetMatch = ""

    lngLongestMatch = lngLongestMatch _
    + Similarity_sub(start1, lngMatchAt1 - 1, _
    start2, lngMatchAt2 - 1, _
    b1, b2, _
    FirstString, _
    strRetMatch1, _
    min_match, _
    recur_level + 1)
    If strRetMatch1 <> "" Then
        RetMatch = RetMatch & strRetMatch1 & "*"
    Else
        RetMatch = RetMatch & IIf(recur_level = 0 _
        And lngLocalLongestMatch > 0 _
        And (lngMatchAt1 > 1 Or lngMatchAt2 > 1) _
        , "*", "")
    End If


    RetMatch = RetMatch & Mid$(FirstString, lngMatchAt1 + 1, lngLocalLongestMatch)


    lngLongestMatch = lngLongestMatch _
    + Similarity_sub(lngMatchAt1 + lngLocalLongestMatch, end1, _
    lngMatchAt2 + lngLocalLongestMatch, end2, _
    b1, b2, _
    FirstString, _
    strRetMatch2, _
    min_match, _
    recur_level + 1)

    If strRetMatch2 <> "" Then
        RetMatch = RetMatch & "*" & strRetMatch2
    Else
        RetMatch = RetMatch & IIf(recur_level = 0 _
        And lngLocalLongestMatch > 0 _
        And ((lngMatchAt1 + lngLocalLongestMatch < end1) _
        Or (lngMatchAt2 + lngLocalLongestMatch < end2)) _
        , "*", "")
    End If

    Similarity_sub = lngLongestMatch

End Function