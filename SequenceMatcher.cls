VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SequenceMatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Ported from Python difflib.SequenceMatcher
'Python Software Foundation License: https://docs.python.org/3/license.html

Dim s1 As String, s2 As String
Dim s1s() As String, s2s() As String
Dim m As Integer, n As Integer
Dim b2j As Scripting.Dictionary
Dim fullbcount As Scripting.Dictionary

Dim longest(2) As Integer
Dim blocks As Collection

Dim opcodes As Collection

Private Sub class_initialize()

End Sub

Public Sub set_seq1(ByVal a As String)

    s1 = a
    m = Len(a)
    ReDim s1s(m)
    
    Set blocks = New Collection
    Set opcodes = New Collection
    
    Dim i As Integer
    For i = 0 To m
        s1s(i) = Mid(s1, i + 1, 1)
    Next i

End Sub

Public Sub set_seq2(ByVal b As String)

    s2 = b
    n = Len(b)
    ReDim s2s(n)
    
    Set blocks = New Collection
    Set opcodes = New Collection
    
    Set b2j = New Scripting.Dictionary
    Set fullbcount = New Scripting.Dictionary
    
    Dim i As Integer
    Dim inj As Collection
    For i = 0 To n
        s2s(i) = Mid(s2, i + 1, 1)
        
        If b2j.Exists(s2s(i)) Then
            b2j(s2s(i)).Add i
        Else
            Set inj = New Collection
            inj.Add i
            b2j.Add s2s(i), inj
        End If
    Next i

End Sub

Public Sub set_seqs(ByVal a As String, b As String)

    Call set_seq1(a)
    Call set_seq2(b)

End Sub

Private Sub find_longest_match(ByVal alo As Integer, ByVal ahi As Integer, ByVal blo As Integer, ByVal bhi As Integer)

    Dim besti As Integer, bestj As Integer, bestsize As Integer
    Dim i As Integer, j As Integer, j_ As Integer, k As Integer
    
    Dim j2len As Scripting.Dictionary, newj2len As Scripting.Dictionary
    Dim b2j_ As Collection
    
    besti = alo
    bestj = blo
    bestsize = 0
    
    Set j2len = New Scripting.Dictionary

    k = 0
    For i = alo To ahi
        If b2j.Exists(s1s(i)) Then
            Set newj2len = New Scripting.Dictionary
            Set b2j_ = b2j(s1s(i))
            For j_ = 1 To b2j_.Count
            
                j = b2j_(j_)
                
                If j >= bhi Then
                    Exit For
                End If
                
                If j < blo Then
                    ' continue
                Else
                
                    If j2len.Exists(j - 1) Then
                        newj2len(j) = j2len(j - 1) + 1
                        k = j2len(j - 1) + 1
                    Else
                        newj2len(j) = 1
                        k = 1
                    End If
                    
                    If k > bestsize Then
                        besti = i - k + 1
                        bestj = j - k + 1
                        bestsize = k
                    End If
                    
                End If
                
            Next j_
            Set j2len = newj2len
        End If
    Next i

    ' VBA ‚Å‚ÍPython‚ÆˆÙ‚È‚èAi ‚ª 1 ‘‚¦‚é‚Ì‚Å’²®
    i = i - 1
    
    If besti > alo And bestj > blo Then
        Do While s1s(besti - 1) = s2s(bestj - 1)
            besti = besti - 1
            bestj = bestj - 1
            bestsize = bestsize + 1
            
            If besti <= alo And bestj <= blo Then
                Exit Do
            End If
        Loop
    End If
    
    If besti + bestsize < ahi And bestj + bestsize < bhi Then
         Do While s1s(besti + bestsize) = s2s(bestj + bestsize)
            bestsize = bestsize + 1
        Loop
    End If
    
    longest(0) = besti
    longest(1) = bestj
    longest(2) = bestsize

End Sub

Private Sub get_matching_blocks()

    Dim la As Integer, lb As Integer
    Dim alo As Integer, ahi As Integer, blo As Integer, bhi As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim queue As Collection
    Dim q(3) As Integer
    Dim matching_blocks As Collection
    
    If blocks.Count > 0 Then
        Exit Sub
    End If
    
    la = m
    lb = n
    
    Set queue = New Collection
    q(0) = 0
    q(1) = m
    q(2) = 0
    q(3) = n
    queue.Add q
    
    Set matching_blocks = New Collection

    Do While queue.Count > 0
        alo = queue(queue.Count)(0)
        ahi = queue(queue.Count)(1)
        blo = queue(queue.Count)(2)
        bhi = queue(queue.Count)(3)
        queue.Remove (queue.Count)
        
        Call find_longest_match(alo, ahi, blo, bhi)
        
        i = longest(0)
        j = longest(1)
        k = longest(2)
        
'       a[alo:i] vs b[blo:j] unknown
'       a[i:i+k] same as b[j:j+k]
'       a[i+k:ahi] vs b[j+k:bhi] unknown
'       if k is 0, there was no matching block
        If k > 0 Then
            matching_blocks.Add (longest)
            If alo < i And blo < j Then
                q(0) = alo
                q(1) = i
                q(2) = blo
                q(3) = j
                queue.Add (q)
            End If
            
            If i + k < ahi And j + k < bhi Then
                q(0) = i + k
                q(1) = ahi
                q(2) = j + k
                q(3) = bhi
                queue.Add (q)
            End If
        End If
    Loop
    
    
'   It's possible that we have adjacent equal blocks in the
'    matching_blocks list now.  Starting with 2.5, this code was added
'   to collapse them.
    Dim i1 As Integer, j1 As Integer, k1 As Integer
    Dim i2 As Integer, j2 As Integer, k2 As Integer
    Dim block As Variant

    Dim non_adjacent As Collection
    Dim x(2) As Integer
    
    Set non_adjacent = New Collection
    
    For Each block In matching_blocks
        i2 = block(0)
        j2 = block(1)
        k2 = block(2)
'        Is this block adjacent to i1, j1, k1?
        If i1 + k1 = i2 And j1 + k1 = j2 Then
'           Yes, so collapse them -- this just increases the length of
'           the first block by the length of the second, and the first
'           block so lengthened remains the block to compare against.
            k1 = k1 + k2
        Else
'           Not adjacent.  Remember the first block (k1==0 means it's
'           the dummy we started with), and make the second block the
'            new block to compare against.
            If k1 Then
                x(0) = i1
                x(1) = j1
                x(2) = k1
                non_adjacent.Add (x)
            End If
            
            i1 = i2
            j1 = j2
            k1 = k2
            
        End If
    Next block
    
    If k1 Then
        x(0) = i1
        x(1) = j1
        x(2) = k1
        non_adjacent.Add (x)
    End If
    
    x(0) = la
    x(1) = lb
    x(2) = 0
    non_adjacent.Add (x)
    
    Set blocks = non_adjacent
    Call sort_blocks
    
  End Sub

Private Sub sort_blocks()

    Dim sorted As Collection
    Dim min As Integer
    Dim i As Integer, j As Integer

    Set sorted = New Collection
    
    Do While blocks.Count > 0
        min = -1
        For i = 1 To blocks.Count
            If min = -1 Then
                min = blocks(i)(0)
                j = i
            ElseIf min > blocks(i)(0) Then
                min = blocks(i)(0)
                j = i
            End If
        Next i
        sorted.Add (blocks(j))
        blocks.Remove (j)
    Loop
    
    Set blocks = sorted

End Sub

Public Function get_opcodes() As Collection

    Dim i As Integer, j As Integer
    Dim ai As Integer, bj As Integer, size As Integer
    Dim block As Variant
    Dim answer(4) As Variant
    Dim answers As Collection
    Dim tag As String
    
    If opcodes.Count > 0 Then
        Set get_opcodes = opcodes
    End If
    
    Call get_matching_blocks
    
    Set answers = New Collection
    
    For Each block In blocks
        ai = block(0)
        bj = block(1)
        size = block(2)
    
'       invariant:  we've pumped out correct diffs to change
'       a[:i] into b[:j], and the next matching block is
'       a[ai:ai+size] == b[bj:bj+size].  So we need to pump
'       out a diff to change a[i:ai] into b[j:bj], pump out
'       the matching block, and move (i,j) beyond the match
        tag = ""
        
        If i < ai And j < bj Then
            tag = "replace"
        ElseIf i < ai Then
            tag = "delete"
        ElseIf j < bj Then
            tag = "insert"
        End If
        
        If tag <> "" Then
            answer(0) = tag
            answer(1) = i
            answer(2) = ai
            answer(3) = j
            answer(4) = bj
            answers.Add answer
        End If
        
        i = ai + size
        j = bj + size
'       the list of matching blocks is terminated by a
'       sentinel with size 0
        If size > 0 Then
            answer(0) = "equal"
            answer(1) = ai
            answer(2) = i
            answer(3) = bj
            answer(4) = j
            answers.Add answer
        End If
    Next block
    
    Set opcodes = answers
    Set get_opcodes = answers
    
End Function

Private Function calculate_ratio(ByVal matches As Integer, ByVal length As Integer) As Single

    If length > 0 Then
        calculate_ratio = 2# * matches / length
        
    Else
        calculate_ratio = 1#
        
    End If


End Function

Public Function ratio() As Double

'Return a measure of the sequences' similarity (float in [0,1]).
'
'Where T is the total number of elements in both sequences, and
'M is the number of matches, this is 2.0*M / T.
'Note that this is 1 if the sequences are identical, and 0 if
'they have nothing in common.
'
'.ratio() is expensive to compute if you haven't already computed
'.get_matching_blocks() or .get_opcodes(), in which case you may
'want to try .quick_ratio() or .real_quick_ratio() first to get an
'upper bound.
'
'>>> s = SequenceMatcher(None, "abcd", "bcde")
'>>> s.ratio()
'0.75
'>>> s.quick_ratio()
'0.75
'>>> s.real_quick_ratio()
'1.0
'

    Dim i As Integer, matches As Integer

    Call get_matching_blocks
    
    matches = 0
    For i = 1 To blocks.Count
        matches = matches + blocks(i)(2)
    Next i
    
    ratio = calculate_ratio(matches, m + n)
            
End Function

Public Function quick_ratio() As Double
'
'Return an upper bound on ratio() relatively quickly.
'
'this isn 't defined beyond that it is an upper bound on .ratio(), and
'is faster to compute.
'
'viewing a and b as multisets, set matches to the cardinality
'of their intersection; this counts the number of matches
'without regard to order, so is clearly an upper bound
    
    Dim i As Integer, matches As Integer, numb As Integer
    Dim elt As String
    Dim avail As Scripting.Dictionary
    
    If fullbcount.Count = 0 Then
        
        For i = 1 To n
            elt = Mid(s2, i, 1)
            If fullbcount.Exists(elt) Then
                fullbcount.Item(elt) = fullbcount.Item(elt) + 1
            Else
                fullbcount.Item(elt) = 1
            End If
        Next i
        
    End If
        
'avail[x] is the number of times x appears in 'b' less the
'number of times we've seen it in 'a' so far ... kinda
    
    Set avail = New Scripting.Dictionary
    matches = 0
    
    For i = 1 To m
        elt = Mid(s1, i, 1)
        If avail.Exists(elt) Then
            numb = avail.Item(elt)
        Else
            If fullbcount.Exists(elt) Then
                numb = fullbcount.Item(elt)
            Else
                numb = 0
            End If
        End If
        
        avail.Item(elt) = numb - 1
        
        If numb > 0 Then
            matches = matches + 1
        End If
    
    Next i
    
    Debug.Print (matches)
    
    quick_ratio = calculate_ratio(matches, m + n)
        
End Function

Public Function real_quick_ratio() As Double

'Return an upper bound on ratio() very quickly.
'
'this isn 't defined beyond that it is an upper bound on .ratio(), and
'is faster to compute than either .ratio() or .quick_ratio().
'

'can't have more matches than the number of elements in the
'shorter sequence

    Dim less As Integer
    
    If m >= n Then
        less = n
    Else
        less = m
    End If

    real_quick_ratio = calculate_ratio(less, m + n)

End Function

