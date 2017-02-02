Sub alibiMourad()
Dim i As Long
i = 1
debut:
Range("A1").Select
Dim Mot As String
Dim Ws As Object
Dim Nbre As Long
Dim Cycle As Long
Dim Trouvé As Variant
Dim CellAddress As Variant
Dim MyValue As String
'col = InputBox("Saisir la couleur *1:Rouge *2:Bleu *3:Jaune *4:Vert *sinon:SansCouleur ", Title:="Color")
'y = InputBox("Saisir type de recherche *1:Pas a Pas *sinon:Auto ", Title:="Color")
'If col = 1 Then
 '   X = 255 'red
  '  Else
   '  If col = 2 Then
    '    X = 15773696 'blue
     '
      '      Else
       '     If col = 3 Then
        '       X = 65535 'Jaune
         '
          '      Else
           '     If col = 4 Then
            '       X = 5287936 'Vert
             '      Else
              '     X = 0
               '
                '   End If
  '               'End If
   '      End If
    '
   ' End If

Mot = Sheets("Feuil2").Range("A" & i).Value
If Mot = "&" Then Exit Sub
'Mot = InputBox("Saisir la valeur à chercher.", Title:="Recherche")
If Mot = "" Then GoTo fin
X = 15773696 'blue
Range("A1").Select
Cycle = 0
'Recherche et arrêt sur les cellules contenant le Mot
For Each Ws In Worksheets
With Ws
.Activate
Set Trouvé = .Cells.Find(What:=Mot, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart)
If Not Trouvé Is Nothing Then
CellAddress = Trouvé.Address
Do
Cycle = Cycle + 1
'Trouvé.Activate
'If Nbre = 1 Then
'MyValue = MsgBox(" La valeur " & Mot & " est enregistrée 1 seule fois ", vbOKOnly, " Message ")
'Exit Sub
'End If
'If Cycle = Nbre Then
'MyValue = MsgBox(" La valeur " & Mot & " sélectionnée est la dernière !", vbOKOnly, "Message")
Sheets("Feuil1").Activate

'Exit Sub
'Else
If y = 1 Then
'MyValue = MsgBox(" La valeur " & Mot & " sélectionnée est la " & Cycle & " emme " & vbLf & _
'" Voulez vous continuer la recherche ? ", vbYesNo, "Message")
If MyValue = vbNo Then Exit For
End If
Set Trouvé = .Cells.FindNext(After:=Trouvé)
'If ActiveCell.Value = Mot Then GoTo allo

ActiveCell.Rows("1:1").EntireRow.Select
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        If X <> 0 Then
        .Color = X
        End If
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
allo:
'End If

Loop While Not Trouvé Is Nothing And Trouvé.Address <> CellAddress
End If
End With
Next Ws
'MyValue = MsgBox(" il y a  " & Cycle & " " & Mot & " dans ce tableau")

X = 5287936 'Vert
Range("A1").Select
Cycle = 0
'Recherche et arrêt sur les cellules contenant le Mot
For Each Ws In Worksheets
With Ws
.Activate
Set Trouvé = .Cells.Find(What:=Mot, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart)
If Not Trouvé Is Nothing Then
CellAddress = Trouvé.Address
Do
Cycle = Cycle + 1
Trouvé.Activate
If Nbre = 1 Then
'MyValue = MsgBox(" La valeur " & Mot & " est enregistrée 1 seule fois ", vbOKOnly, " Message ")
Exit Sub
End If
If Cycle = Nbre Then
'MyValue = MsgBox(" La valeur " & Mot & " sélectionnée est la dernière !", vbOKOnly, "Message")
Sheets("Feuil1").Activate

Exit Sub
Else
If y = 1 Then
'MyValue = MsgBox(" La valeur " & Mot & " sélectionnée est la " & Cycle & " emme " & vbLf & _
'" Voulez vous continuer la recherche ? ", vbYesNo, "Message")
If MyValue = vbNo Then Exit For
End If
Set Trouvé = .Cells.FindNext(After:=Trouvé)
ActiveCell.Select
'If ActiveCell.Value = Mot Then GoTo alla
With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        If X <> 0 Then
        .Color = X
        End If
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
alla:
End If

Loop While Not Trouvé Is Nothing And Trouvé.Address <> CellAddress
End If
End With
Next Ws
'MyValue = MsgBox(" il y a  " & Cycle & " " & Mot & " dans ce tableau")
fin:
i = i + 1
GoTo debut

End Sub
