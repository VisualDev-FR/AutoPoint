Attribute VB_Name = "M01_CSV"
Option Explicit

Private Const csvFilePath As String = "C:\Users\menan\source\repos\AutoPoint\AutoPoint.csv"
Private csvFormat As Tristate

Public Sub WriteCSV()

    csvFormat = TristateTrue

    Dim tPointages As Variant: tPointages = Range("Tab_Pointages").Value

    Dim fso As New FileSystemObject
    Dim mStream As TextStream
    Dim strPointage() As String
    Dim i As Integer
    
    fso.GetFile(csvFilePath).Copy Destination:=Environ("AppData") & "\VBA\AutoPoint\AutoPoint_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".csv"
    
    'Lecture du fichier csv

    Set mStream = fso.OpenTextFile(csvFilePath, ForWriting, False, csvFormat)
    
    For i = LBound(tPointages, 1) To UBound(tPointages, 1)
        
        ReDim strPointage(0 To 5)

        strPointage(0) = tPointages(i, 4) 'Projet
        strPointage(1) = tPointages(i, 5) 'Tache
        strPointage(2) = tPointages(i, 6) 'ssTache
        
        strPointage(3) = Format(CDate(tPointages(i, 1)), "dd/mm/yyyy") 'Date
        strPointage(4) = Format(CDate(tPointages(i, 2)), "hh:mm") 'Debut
        strPointage(5) = Format(CDate(tPointages(i, 3)), "hh:mm") 'Fin

        mStream.WriteLine Join(strPointage, ";")

    Next i
    
    mStream.Close

End Sub

Public Sub ReadCSV()

    csvFormat = TristateTrue

    Dim fso As New FileSystemObject
    Dim mStream As TextStream
    Dim strPointage() As String, dicKey As String
    
    Dim mClosePoint As Boolean
    Dim closetime As String

    Dim importedPoints As Variant
    Dim localPoints As Variant
    
    Dim localDico As New Dictionary
    Dim importedDico As New Dictionary
    Dim dicoPointage As Dictionary
    
    Dim i As Integer
    
    fso.GetFile(csvFilePath).Copy Destination:=Environ("AppData") & "\VBA\AutoPoint\AutoPoint_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".csv"
    
    'Lecture du fichier csv

    Set mStream = fso.OpenTextFile(csvFilePath, ForReading, False, csvFormat)
    
    With mStream
    
        importedPoints = Split(.ReadAll, vbCrLf)
        .Close
        
    End With
    
    'Traitement des pointages issus du ficher csv
    
    For i = LBound(importedPoints, 1) To UBound(importedPoints, 1)
    
        If importedPoints(i) <> "" Then
        
            strPointage = Split(importedPoints(i), ";")
            
            mClosePoint = UBound(strPointage, 1) > 4
            
            If mClosePoint Then
                closetime = strPointage(5)
            Else
                closetime = ""
            End If
            
            dicKey = strPointage(3) & strPointage(4) & closetime
    
            If Not importedDico.Exists(dicKey) Then
                
                Set dicoPointage = New Dictionary

                dicoPointage.Add Key:="PROJET", Item:=strPointage(0)
                dicoPointage.Add Key:="TACHE", Item:=strPointage(1)
                dicoPointage.Add Key:="SSTACHE", Item:=strPointage(2)
                dicoPointage.Add Key:="DATE", Item:=strPointage(3)
                dicoPointage.Add Key:="OPEN", Item:=strPointage(4)
                dicoPointage.Add Key:="CLOSE", Item:=closetime
                
                importedDico.Add Key:=dicKey, Item:=dicoPointage
                
            End If
        
        End If

    Next i

    'Traitement des pointages issus du tableau local
    
    localPoints = Range("Tab_Pointages").Value
    
    For i = LBound(localPoints, 1) To UBound(localPoints, 1)
    
        If localPoints(i, 1) <> "" Then

            mClosePoint = localPoints(i, 3) <> ""
            
            dicKey = localPoints(i, 1) & Format(localPoints(i, 2), "hh:mm") & Format(IIf(mClosePoint, localPoints(i, 3), ""), "hh:mm")
    
            If Not localDico.Exists(dicKey) Then
                
                Set dicoPointage = New Dictionary

                dicoPointage.Add Key:="PROJET", Item:=localPoints(i, 4)
                dicoPointage.Add Key:="TACHE", Item:=localPoints(i, 5)
                dicoPointage.Add Key:="SSTACHE", Item:=localPoints(i, 6)
                dicoPointage.Add Key:="DATE", Item:=localPoints(i, 1)
                dicoPointage.Add Key:="OPEN", Item:=Format(localPoints(i, 2), "hh:mm")
                dicoPointage.Add Key:="CLOSE", Item:=Format(IIf(mClosePoint, localPoints(i, 3), ""), "hh:mm")
                
                localDico.Add Key:=dicKey, Item:=dicoPointage
                
            End If
        
        End If

    Next i
    
    'Import des nouveaux pointages
    
    Dim k
    
    For Each k In importedDico.Keys
    
        If Not localDico.Exists(k) Then localDico.Add Key:=k, Item:=importedDico(k)

    Next
    
    ReDim importedPoints(1 To localDico.Count, 1 To UBound(localPoints, 2) - 1)
    
    i = 0
    
    For Each k In localDico.Keys
        
        i = i + 1

        On Error Resume Next
        
        With localDico(k)

            importedPoints(i, 1) = CDbl(CDate(.Item("DATE")))
            importedPoints(i, 2) = CDate(.Item("OPEN"))
            importedPoints(i, 3) = CDate(.Item("CLOSE"))

            importedPoints(i, 4) = .Item("PROJET")
            importedPoints(i, 5) = .Item("TACHE")
            importedPoints(i, 6) = .Item("SSTACHE")

        End With
        
        On Error GoTo -1
        
    Next
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    With Range("Tab_Pointages")
   
        .Resize(UBound(importedPoints, 1), UBound(importedPoints, 2)) = importedPoints

    End With
    
    Range("Tab_Pointages[Date]").NumberFormat = "dd/mm/yyyy"
    
    Range("Tab_Pointages[[Date]:[Fin]]").Value = Range("Tab_Pointages[[Date]:[Fin]]").Value

    With ThisWorkbook
        .RefreshAll
        .Save
    End With
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
