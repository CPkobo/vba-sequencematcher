Attribute VB_Name = "WordSeqApplyer"
Option Explicit

Dim seq As SequenceMatcher
Dim src As String, tgt As String
Dim ccs As ContentControls
Dim cc As ContentControl

Sub main()

    Call create_seqence_matcher
    
    'Debug.Print seq.ratio
    'Debug.Print seq.quick_ratio
    'Debug.Print seq.real_quick_ratio

    'Call apply_opcode

End Sub

Private Sub create_seqence_matcher()

    Set seq = New SequenceMatcher

    Set ccs = ThisDocument.ContentControls
    
    For Each cc In ccs
    
        If cc.tag = "src" Then
            src = cc.Range.Text
        ElseIf cc.tag = "tgt" Then
            tgt = cc.Range.Text
        ElseIf cc.tag = "res" Then
            cc.Range.Text = src
            cc.Range.Select
        End If
    
    Next cc
    
    Call seq.set_seqs(src, tgt)

End Sub

Public Sub apply_opcode()

    Dim opcodes As Collection
    Dim rng As Range, rng2 As Range
    Dim i As Integer
    Dim code() As Variant
    Dim isRevMode As Boolean: isRevMode = ThisDocument.TrackRevisions
    Dim insText As String
    Dim selStart As Long, selEnd As Long
    
    Call initialize
    
    Set opcodes = seq.get_opcodes
        
    Set rng = Selection.Range
    selStart = rng.Start
    selEnd = rng.End
    
    ThisDocument.TrackRevisions = True
    
    For i = opcodes.Count To 1 Step -1
    
        code = opcodes(i)
        
        Set rng2 = rng
        
        If code(0) = "equal" Then
        
        ElseIf code(0) = "insert" Then
            rng2.SetRange Start:=selStart + code(1), End:=selStart + code(1)
            rng2.Select
            insText = Mid(tgt, code(3) + 1, code(4) - code(3))
            Selection.TypeText insText
            
        ElseIf code(0) = "delete" Then
            rng2.SetRange Start:=selStart + code(1), End:=selStart + code(2)
            rng2.Select
            Selection.Delete
        
        ElseIf code(0) = "replace" Then
            rng2.SetRange Start:=selStart + code(1), End:=selStart + code(2)
            rng2.Select
            Selection.Delete
            rng2.SetRange Start:=selStart + code(1), End:=selStart + code(1)
            rng2.Select
            insText = Mid(tgt, code(3) + 1, code(4) - code(3))
            Selection.TypeText insText
        
        End If
    
    Next i
    
    ThisDocument.TrackRevisions = isRevMode

End Sub

