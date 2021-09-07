Attribute VB_Name = "Module1"
Option Explicit

Global Ply As Byte '��� ������
Global SaveEvalWin As Integer '���� EvalWin
Global PLine(0 To 84) As Byte
Global Depth As Byte '��� ������
Dim QStart As Long, QFinish As Long '����� ����� ������
Dim Best As Integer '���� ���� ������

'����� �����
Public Enum enumColor
    CEmpty = 0
    Cwhite = 1
    Cblack = 2
End Enum


'���� ������ �� ����� �� ���� ������
Public Type TMove
    Target  As Byte '��� ������
    File    As Byte '���� ������
    Rank    As Byte '�� ������
    Pcolor  As enumColor '��� ������
    PPlayed As Boolean '�� ������ ������ �� ��
    
End Type
Global Moves(1 To 42) As TMove '������ ���� �������
Global LegalMoves(1 To 7) As Byte '������ ������� �������
Global NumMoves As Byte '��� ������� �������
'�� ��� ������ ��� ��� ����
Global WhiteToMove As Boolean
Global PlayMethod As Byte '����� ����� �� �� ���� ���� �� ��� ���� �� ������
Global SetupMode As Boolean '�� �� ��� setup
Global LastColo As enumColor '��� ��� �� ���� setup
Global Nodes As Long '��� ����� ��������
Global QNodes As Long '��� ����� ��������
Global StopFindWin As Boolean '�� ���� ����� �� ������ �������
Sub OrderMoves(Moves() As Byte, NumMoves As Byte)
'����� ��� ����� �������
'���� ������ alpha ,beta cuttoff
Dim i As Byte
Dim Values(1 To 7) As Byte
For i = 1 To NumMoves
    Select Case Moves(i)
     Case 3, 4, 5, 10, 11, 12, 17, 18, 19, 24, 25, 26
     Values(i) = Values(i) + 20
     Case 1, 2, 6, 7, 9, 13
     Values(i) = Values(i) + 15
     Case 31, 32, 33, 38, 39, 40
     Values(i) = Values(i) + 10
    End Select
Next
QSortMoves Moves(), Values(), 1, NumMoves

End Sub
Private Sub QSortMoves(Moves() As Byte, Values() As Byte, ByVal iStart As Byte, ByVal iEnd As Byte)
'����� ������� ����� �����
Dim Partition   As Byte, TempValue    As Byte
Dim i           As Byte, j As Byte
Dim TempMove    As Byte

If iEnd > iStart Then
    i = iStart
    j = iEnd
    Partition = Values((i + j) / 2)
    Do
        Do While Values(i) > Partition
            i = i + 1
        Loop
        Do While Values(j) < Partition
            j = j - 1
        Loop
        If i <= j Then
            TempValue = Values(i)
            Values(i) = Values(j)
            Values(j) = TempValue
            TempMove = Moves(i)
            Moves(i) = Moves(j)
            Moves(j) = TempMove
            
            i = i + 1
            j = j - 1
        End If
    Loop While i <= j
    QSortMoves Moves(), Values(), i, iEnd
    QSortMoves Moves(), Values(), iStart, j

End If

End Sub

Public Function EvalWin() As Integer
Dim a As Byte
EvalWin = 0

For a = 1 To 39 '��� �� ���� ������ �� ������ 39
'����� 10000 �� ��� ���� �� ���� ����
If Moves(a).Pcolor = Cwhite Then
    If (Moves(a + 1).Pcolor = Cwhite) And (Moves(a + 1).Rank = Moves(a).Rank) _
    And (Moves(a + 2).Pcolor = Cwhite) And (Moves(a + 2).Rank = Moves(a).Rank) _
    And (Moves(a + 3).Pcolor = Cwhite) And (Moves(a + 3).Rank = Moves(a).Rank) Then EvalWin = 10000 - Ply: Exit Function
End If
'����� 10000�����
If Moves(a).Pcolor = Cblack Then
    If (Moves(a + 1).Pcolor = Cblack) And (Moves(a + 1).Rank = Moves(a).Rank) _
    And (Moves(a + 2).Pcolor = Cblack) And (Moves(a + 2).Rank = Moves(a).Rank) _
    And (Moves(a + 3).Pcolor = Cblack) And (Moves(a + 3).Rank = Moves(a).Rank) Then EvalWin = -10000 + Ply: Exit Function
End If
Next

For a = 1 To 21 '��� ���� ���� ������ �� ������ 21
'����� 10000 �� ��� ���� ���� ���� ����
If Moves(a).Pcolor = Cwhite Then
    If (Moves(a + 7).Pcolor = Cwhite) And (Moves(a + 7).File = Moves(a).File) _
    And (Moves(a + 14).Pcolor = Cwhite) And (Moves(a + 14).File = Moves(a).File) _
    And (Moves(a + 21).Pcolor = Cwhite) And (Moves(a + 21).File = Moves(a).File) Then EvalWin = 10000 - Ply: Exit Function
End If
'����� 10000�����
If Moves(a).Pcolor = Cblack Then
    If (Moves(a + 7).Pcolor = Cblack) And (Moves(a + 7).File = Moves(a).File) _
    And (Moves(a + 14).Pcolor = Cblack) And (Moves(a + 14).File = Moves(a).File) _
    And (Moves(a + 21).Pcolor = Cblack) And (Moves(a + 21).File = Moves(a).File) Then EvalWin = -10000 + Ply: Exit Function
End If
Next

For a = 1 To 18 '��� ��� ����� ���� ������ �� ������ 18
'����� 10000 �� ��� ���� ��� ����� ����
If Moves(a).Pcolor = Cwhite Then
    If (Moves(a + 8).Pcolor = Cwhite) And (Moves(a + 8).File = Moves(a).File + 1) And (Moves(a + 8).Rank = Moves(a).Rank + 1) _
    And (Moves(a + 16).Pcolor = Cwhite) And (Moves(a + 16).File = Moves(a).File + 2) And (Moves(a + 16).Rank = Moves(a).Rank + 2) _
    And (Moves(a + 24).Pcolor = Cwhite) And (Moves(a + 24).File = Moves(a).File + 3) And (Moves(a + 24).Rank = Moves(a).Rank + 3) Then EvalWin = 10000 - Ply: Exit Function
End If

If Moves(a).Pcolor = Cblack Then
    If (Moves(a + 8).Pcolor = Cblack) And (Moves(a + 8).File = Moves(a).File + 1) And (Moves(a + 8).Rank = Moves(a).Rank + 1) _
    And (Moves(a + 16).Pcolor = Cblack) And (Moves(a + 16).File = Moves(a).File + 2) And (Moves(a + 16).Rank = Moves(a).Rank + 2) _
    And (Moves(a + 24).Pcolor = Cblack) And (Moves(a + 24).File = Moves(a).File + 3) And (Moves(a + 24).Rank = Moves(a).Rank + 3) Then EvalWin = -10000 + Ply: Exit Function
End If
Next

For a = 22 To 39 '��� ��� ����� ���� ������ �� ������ 22 ������ �� 39
'����� 10000 �� ��� ���� ��� ����� ����
If Moves(a).Pcolor = Cwhite Then
    If (Moves(a - 6).Pcolor = Cwhite) And (Moves(a - 6).File = Moves(a).File + 1) And (Moves(a - 6).Rank = Moves(a).Rank - 1) _
    And (Moves(a - 12).Pcolor = Cwhite) And (Moves(a - 12).File = Moves(a).File + 2) And (Moves(a - 12).Rank = Moves(a).Rank - 2) _
    And (Moves(a - 18).Pcolor = Cwhite) And (Moves(a - 18).File = Moves(a).File + 3) And (Moves(a - 18).Rank = Moves(a).Rank - 3) Then EvalWin = 10000 - Ply: Exit Function
End If
If Moves(a).Pcolor = Cblack Then
    If (Moves(a - 6).Pcolor = Cblack) And (Moves(a - 6).File = Moves(a).File + 1) And (Moves(a - 6).Rank = Moves(a).Rank - 1) _
    And (Moves(a - 12).Pcolor = Cblack) And (Moves(a - 12).File = Moves(a).File + 2) And (Moves(a - 12).Rank = Moves(a).Rank - 2) _
    And (Moves(a - 18).Pcolor = Cblack) And (Moves(a - 18).File = Moves(a).File + 3) And (Moves(a - 18).Rank = Moves(a).Rank - 3) Then EvalWin = -10000 + Ply: Exit Function
End If
Next

End Function
Sub MakeMove(num As Byte, colo As enumColor)
Ply = Ply + 1
Moves(num).PPlayed = True
Moves(num).Pcolor = colo
WhiteToMove = Not WhiteToMove

End Sub
Sub UnMakeMove(num As Byte)
Ply = Ply - 1
Moves(num).PPlayed = False
Moves(num).Pcolor = CEmpty
WhiteToMove = Not WhiteToMove

End Sub
Public Function Search(ByVal Depth As Integer) As Integer

Dim Score As Integer, NewScore As Integer
Dim LegalMovesNow(1 To 7) As Byte '���� ������ ������� �������� ��� ��� �������
Dim i As Byte, b As Integer
Dim colo As enumColor

SaveEvalWin = EvalWin
If SaveEvalWin <> 0 Then
    If WhiteToMove Then Search = SaveEvalWin: Exit Function Else Search = -SaveEvalWin: Exit Function
End If

Erase LegalMovesNow()
Score = -11000

If Depth = 0 Then
    Score = Eval
    Search = Score
    Exit Function
End If
GenerateMoves
For i = 1 To NumMoves: LegalMovesNow(i) = LegalMoves(i): Next

For i = 1 To NumMoves
 '����� ��� ������
 If WhiteToMove Then colo = Cwhite Else colo = Cblack
 
 Call MakeMove(LegalMovesNow(i), colo)
 Nodes = Nodes + 1
 NewScore = -Search(Depth - 1)

 Call UnMakeMove(LegalMovesNow(i))
 If NewScore > Score Then Score = NewScore: PLine(Ply) = LegalMovesNow(i)
 Next

Search = Score

End Function
Public Function QSearch(ByVal Depth As Integer, alpha As Integer, beta As Integer) As Integer

Dim NewScore As Integer
Dim LegalMovesNow(1 To 7) As Byte '���� ������ ������� �������� ��� ��� �������
Dim i As Byte, c As Integer
Dim colo As enumColor


SaveEvalWin = EvalWin
If SaveEvalWin <> 0 Then
    If WhiteToMove Then QSearch = SaveEvalWin: Exit Function Else QSearch = -SaveEvalWin: Exit Function
End If


Erase LegalMovesNow()
If Depth = 0 Then
    alpha = Eval
    QSearch = alpha
    Exit Function
End If
GenerateMoves
OrderMoves LegalMoves(), NumMoves
For i = 1 To NumMoves: LegalMovesNow(i) = LegalMoves(i): Next

For i = 1 To NumMoves
 '����� ��� ������
 If WhiteToMove Then colo = Cwhite Else colo = Cblack
 
 Call MakeMove(LegalMovesNow(i), colo)
  QNodes = QNodes + 1
 
 NewScore = -QSearch(Depth - 1, -beta, -alpha)
  
 Call UnMakeMove(LegalMovesNow(i))
 

 If NewScore >= beta Then QSearch = beta: PLine(Ply) = LegalMovesNow(i): Exit Function
 If NewScore > alpha Then alpha = NewScore: PLine(Ply) = LegalMovesNow(i)

  Next

QSearch = alpha


End Function

Public Function Eval() As Integer '�� ���� ��� ������
Dim Score As Integer '����� �������
Dim i As Byte '����
Dim b As Byte
Score = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���� ���� ������� ������� ������
For i = 1 To 42

  If Moves(i).Pcolor = Cwhite Then
   Select Case Moves(i).Target
     Case 3, 4, 5
     Score = Score + 20
     Case 10, 11, 12
     Score = Score + 22
     Case 17, 18, 19, 24, 25, 26
     Score = Score + 25
     Case 2, 6, 9, 13, 16, 23, 20, 27, 31, 32, 33, 38, 39, 40, 1, 7
     Score = Score + 15
     Case 8, 15, 22, 14, 21, 28
     Score = Score + 2
    End Select
  End If
  
  If Moves(i).Pcolor = Cblack Then
   Select Case Moves(i).Target
     Case 3, 4, 5
     Score = Score - 20
     Case 10, 11, 12
     Score = Score - 22
     Case 17, 18, 19, 24, 25, 26
     Score = Score - 25
     Case 2, 6, 9, 13, 16, 23, 20, 27, 31, 32, 33, 38, 39, 40, 1, 7
     Score = Score - 15
     Case 8, 15, 22, 14, 21, 28
     Score = Score - 2
  End Select
  End If
   
  Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Not WhiteToMove Then
    Eval = -Score
    Else
    Eval = Score
End If

End Function

Public Function GenerateMoves() As Byte '��� ������� �������

Dim i As Byte '����
Dim b As Byte '����


Erase LegalMoves()

NumMoves = 0

If SaveEvalWin <> 0 Then Exit Function



 For b = 1 To 7
    i = 0
200     If Moves(b + i).PPlayed = False Then
         NumMoves = NumMoves + 1: LegalMoves(NumMoves) = b + i
    Else: i = i + 7: If i < 36 Then GoTo 200
    End If
Next

DoEvents
GenerateMoves = NumMoves


End Function
