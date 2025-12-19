Attribute VB_Name = "Marker"
Option Explicit
Public ColorCode As Long


Sub QuickAnalysis()

    Dim i As Integer, LC As Integer
    Dim Batch As String
    
        LC = Tabelle1.Cells(4, 16384).End(xlToLeft).Column
        
        For i = 5 To LC
            Batch = Tabelle1.Cells(4, i).value
            
            If Tabelle1.Cells(10, i).value = "n.i.O." Then
                ColorCode = 255
                Call MultiMarker(Batch, True)
            End If
        Next i
        
'Optionaler Part. Vorerst auskommentiert
'        For i = 5 To LC
'            Batch = Tabelle1.Cells(4, i).value
'
'            If Tabelle1.Cells(10, i).value = "b.i.O." Then
'                ColorCode = 65535
'                Call MultiMarker(Batch, True)
'            End If
'        Next i
            
        For i = 5 To LC
            Batch = Tabelle1.Cells(4, i).value

            If Tabelle1.Cells(10, i).value = "i.O." Then
                ColorCode = 32768
                Call MultiMarker(Batch, True)
            End If
        Next i
        
        For i = 5 To LC
            Batch = Tabelle1.Cells(4, i).value

            If Tabelle1.Cells(10, i).value = "" Then
                ColorCode = 0
                Call MultiMarker(Batch, True)
            End If
        Next i

End Sub


Function GetColor()

    If Application.Dialogs(xlDialogEditColor).Show(1, 0, 0, 0) = True Then
        GetColor = ActiveWorkbook.Colors(1)
    Else
        'Nichts tun, falls Cancel gedrückt wurde
        End
    End If
    
End Function


Sub CheckAll()
    
    Dim i As Integer, j As Integer, k As Integer, LC As Integer, LR As Integer
    Dim SearchTerm As Variant
    
        With ActiveSheet
            If MsgBox("Achtung. Hierdurch werden alle vorhandenen Formatierungen zurückgesetzt" & vbCrLf & vbCrLf & "Fortfahren?", vbYesNo, "") = vbNo Then
                Exit Sub
            Else
                ColorCode = GetColor
                Call Reset_Table

                LC = .Cells(4, 16384).End(xlToLeft).Column
                LR = .Cells(1048576, 3).End(xlUp).Row
                    
                For k = 5 To LC
                    For i = 13 To LR
                        If .Cells(i, 2).value = "Rohstoff" Then
                            SearchTerm = Split(.Cells(i, k), "|")
                            For j = LBound(SearchTerm) To UBound(SearchTerm)
                                If MultiChecker(ReplaceBetweenBrackets(SearchTerm(j), "")) = (LC - 4) Then Mark_Cells ReplaceBetweenBrackets(SearchTerm(j), ""), ColorCode
                            Next j
                        End If
                    Next i
                Next k
                
                MsgBox "Überprüfung abgeschlossen. Rohstoffe, die sich in allen " & (LC - 4) & " Chargen befinden, wurden markiert.", vbOKOnly, "Markierung beendet"
            End If
        End With
End Sub

Sub Reset_Table()
    
    Dim LR As Integer, LC As Integer, i As Integer, j As Integer
    
        With ActiveSheet
            LR = .Cells(1048576, 3).End(xlUp).Row
            LC = .Cells(4, 16384).End(xlToLeft).Column
            
            For i = 13 To LR
                For j = 5 To LC
                    .Cells(i, j).Font.ColorIndex = xlAutomatic
                    .Cells(i, j).Font.TintAndShade = 0
                    .Cells(i, j).Font.Bold = False
                    
                    .Cells(i, j).Interior.Pattern = xlNone
                    .Cells(i, j).Interior.TintAndShade = 0
                    .Cells(i, j).Interior.PatternTintAndShade = 0
                Next j
            Next i
        End With
End Sub
Sub Mark_Cells(ByVal SearchTerm As String, ColorCode As Long)
    
    Dim SearchString As String
    Dim LR As Integer, LC As Integer, i As Integer, j As Integer
    Dim SearchIndex As Long
        
        With ActiveSheet
            If SearchTerm = "" Then SearchTerm = .Cells(1, 3).value
            LR = .Cells(1048576, 3).End(xlUp).Row
            LC = .Cells(4, 16384).End(xlToLeft).Column
            SearchString = Replace(Replace(SearchTerm, vbCrLf, ""), Chr(10), "")
            Do Until Left(SearchString, 1) <> " "
                SearchString = Right(SearchString, Len(SearchString) - 1)
            Loop
            
            For i = 5 To LC
                For j = 13 To LR
                    SearchIndex = 1
                    
                    Do While SearchIndex > 0
                        SearchIndex = InStr(SearchIndex, CStr(.Cells(j, i).value), SearchString)
                        If SearchIndex > 0 Then
                            .Cells(j, i).Characters(Start:=SearchIndex, Length:=Len(SearchString)).Font.Color = ColorCode
                            SearchIndex = SearchIndex + Len(SearchString)
                        End If
                    Loop
                Next j
            Next i
        End With
End Sub

Function ReplaceBetweenBrackets(ByVal originalText As String, ByVal replacementText As String) As String
    
    Dim startIdx As Long, EndIdx As Long
        
        ' Finden Sie die Position des ersten " ("
        startIdx = InStr(originalText, " (")
        
        ' Finden Sie die Position des letzten ")"
        EndIdx = InStrRev(originalText, ")")
        
        If startIdx > 0 And EndIdx > 0 And startIdx < EndIdx Then
            ' Extrahieren Sie den Text zwischen " (" und ")"
            Dim textBetween As String
            textBetween = Mid(originalText, startIdx, EndIdx - startIdx + 1)
            
            ' Ersetzen Sie den gefundenen Text
            textBetween = Replace(originalText, textBetween, replacementText)
            ReplaceBetweenBrackets = textBetween
        Else
            ' Wenn keine " (" oder ")" gefunden wurden, geben Sie den ursprünglichen Text zurück
            ReplaceBetweenBrackets = originalText
        End If
End Function

Function MultiChecker(ByVal SuchBegr As String)
    
    Dim i As Integer, j As Integer, LC As Integer, LR As Integer, Checker As Integer
    
        With ActiveSheet
            LC = .Cells(4, 16384).End(xlToLeft).Column
            LR = .Cells(1048576, 3).End(xlUp).Row
            
            SuchBegr = Replace(Replace(SuchBegr, vbCrLf, ""), Chr(10), "")
            Do Until Left(SuchBegr, 1) <> " "
                SuchBegr = Right(SuchBegr, Len(SuchBegr) - 1)
            Loop
            
            For i = 13 To LR
                For j = 5 To LC
                    If InStr(.Cells(i, j).value, SuchBegr) Then Checker = Checker + 1
                Next j
            Next i
        End With
        
        MultiChecker = Checker
End Function
Sub MultiMarker(ByVal Batch As String, Optional QuickAnalysis As Boolean = False)
    
    Dim i As Integer, j As Integer, k As Integer, LC As Integer, LR As Integer
    Dim SearchTerm As Variant
    
        With ActiveSheet
            LC = .Cells(4, 16384).End(xlToLeft).Column
            LR = .Cells(1048576, 3).End(xlUp).Row
            
            For i = 4 To LC
                If CStr(.Cells(4, i).value) = Batch Then
                    k = i
                    Exit For
                End If
            Next i
            If i > LC Then MsgBox "Die gewählte Charge wurde in der Tabelle nicht gefunden", vbOKOnly, "Fehler": Exit Sub
            
            For i = 13 To LR
                If .Cells(i, 2).value = "Rohstoff" Then
                    SearchTerm = Split(.Cells(i, k), "|")
                    For j = LBound(SearchTerm) To UBound(SearchTerm)
                        If QuickAnalysis = True Then
                            Mark_Cells ReplaceBetweenBrackets(SearchTerm(j), ""), ColorCode
                        Else
                            If MultiChecker(ReplaceBetweenBrackets(SearchTerm(j), "")) > 1 Then Mark_Cells ReplaceBetweenBrackets(SearchTerm(j), ""), ColorCode
                        End If
                    Next j
                End If
            Next i
            
        End With
End Sub

