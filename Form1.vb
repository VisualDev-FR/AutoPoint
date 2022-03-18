Imports System.IO

Public Class Form1

    Private Const fileCsvPath As String = "C:\Users\menan\source\repos\AutoPoint\AutoPoint.csv"

    Public csvEncoding As System.Text.Encoding = System.Text.Encoding.Unicode

    Private lastPointage As Pointage
    Private dicoProjet As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))

    Public lastPoint As String
    '---------------------------------------------------------------------------------------------------------------------------------------
    'GESTION DES EVENEMENTS
    '---------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim lastCsvLine As String

        lastPointage = New Pointage

        lastCsvLine = GetLastPoint()

        lastPointage.Read(lastCsvLine)

        If lastPointage.Status = "Open" Then
            btn_OpenPoint.Text = "Fermer Pointage"
            txt_Projet.Enabled = False
            txt_Tache.Enabled = False
            txt_SSTache.Enabled = False
        Else
            btn_OpenPoint.Text = "Ouvrir Pointage"
        End If

        Call WriteLog()
        Call InitTaskLists()

        Timer1.Start()

    End Sub

    Private Sub btn_OpenPoint_Click(sender As Object, e As EventArgs) Handles btn_OpenPoint.Click

        If lastPointage Is Nothing Then Exit Sub

        Select Case lastPointage.Status

            Case "Closed"

                If txt_Projet.Text = "" Or txt_Tache.Text = "" Or txt_SSTache.Text = "" Then

                    MessageBox.Show("Veuillez renseigner un Projet + une tâche + une sous-tâche.")
                    Exit Sub

                End If

                btn_OpenPoint.Text = "Fermer Pointage"

                lastPointage = New Pointage
                lastPointage.Open(txt_Projet.Text, txt_Tache.Text, txt_SSTache.Text)

                txt_Projet.Enabled = False
                txt_Tache.Enabled = False
                txt_SSTache.Enabled = False

                Call InitTaskLists()

                Me.WindowState = FormWindowState.Minimized

            Case "Open"

                btn_OpenPoint.Text = "Ouvrir Pointage"
                lastPointage.Close()
                Me.Text = "AutoPoint"
                txt_Projet.Enabled = True
                txt_Tache.Enabled = True
                txt_SSTache.Enabled = True

        End Select

        Call WriteLog()

    End Sub
    Private Sub Timer1_Tick() Handles Timer1.Tick 'sender As Object, e As EventArgs

        With lastPointage

            If .Status = "Open" Then

                Dim openDate As DateTime = Convert.ToDateTime(.Open_Hour)
                Dim lapsDate As TimeSpan = Now - openDate

                .Duration = Format(lapsDate.Hours, "00:") & Format(lapsDate.Minutes, "00:") & Format(lapsDate.Seconds, "00")

                If Me.WindowState = FormWindowState.Minimized Then Me.Text = .Duration

                Call WriteLog()

            End If

        End With

    End Sub

    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged

        If Me.WindowState = FormWindowState.Minimized Then
            Me.Text = lastPointage.Duration
        Else
            Me.Text = "AutoPoint"
        End If

    End Sub

    Private Sub txt_Tache_Validated(sender As Object, e As EventArgs) Handles txt_Tache.Validated
        Call InitTaskLists()
    End Sub

    Private Sub txt_Projet_Validated(sender As Object, e As EventArgs) Handles txt_Projet.Validated
        Call InitTaskLists()
    End Sub

    Private Sub txt_Tache_SelectedValueChanged(sender As Object, e As EventArgs) Handles txt_Tache.SelectedValueChanged
        txt_SSTache.Text = ""
    End Sub

    Private Sub txt_Projet_SelectedValueChanged(sender As Object, e As EventArgs) Handles txt_Projet.SelectedValueChanged
        txt_Tache.Text = ""
        txt_SSTache.Text = ""
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------
    'PROCEDURES
    '---------------------------------------------------------------------------------------------------------------------------------------

    Private Sub WriteLog()

        txt_Projet.Text = lastPointage.Projet
        txt_Tache.Text = lastPointage.Tache
        txt_SSTache.Text = lastPointage.SousTache

        logFrame.Text = lastPointage.strLog

    End Sub
    Private Function GetLastPoint() As String

        Dim mStreamReader As StreamReader
        Dim tablePointFromCsv() As String

        mStreamReader = My.Computer.FileSystem.OpenTextFileReader(fileCsvPath, csvEncoding)

        tablePointFromCsv = Split(mStreamReader.ReadToEnd, vbCrLf)

        mStreamReader.Close()

        If tablePointFromCsv(UBound(tablePointFromCsv)) = "" Then
            Return tablePointFromCsv(UBound(tablePointFromCsv) - 1)
        Else
            Return tablePointFromCsv(UBound(tablePointFromCsv))
        End If

    End Function
    Private Sub InitTaskLists()

        Dim dicoTache As New Dictionary(Of String, Dictionary(Of String, String))
        Dim dicoSSTache As New Dictionary(Of String, String)
        Dim pointsTable As String()
        Dim i As Integer

        Dim mStreamReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(fileCsvPath, csvEncoding)

        pointsTable = Split(mStreamReader.ReadToEnd, vbCrLf)

        mStreamReader.Close()

        For i = LBound(pointsTable, 1) To UBound(pointsTable, 1)

            If pointsTable(i) <> "" Then

                Dim activeProjet As String = Split(pointsTable(i), ";")(0)
                Dim activeTache As String = Split(pointsTable(i), ";")(1)
                Dim activeSSTache As String = Split(pointsTable(i), ";")(2)

                If Not dicoProjet.ContainsKey(activeProjet) Then dicoProjet.Add(activeProjet, New Dictionary(Of String, Dictionary(Of String, String)))

                If Not dicoProjet(activeProjet).ContainsKey(activeTache) Then dicoProjet(activeProjet).Add(activeTache, New Dictionary(Of String, String))

                If Not dicoProjet(activeProjet)(activeTache).ContainsKey(activeSSTache) Then dicoProjet(activeProjet)(activeTache).Add(activeSSTache, activeSSTache)

            End If

        Next i

        txt_Projet.Items.Clear()
        txt_Tache.Items.Clear()
        txt_SSTache.Items.Clear()

        For Each strProjet In dicoProjet.Keys

            txt_Projet.Items.Add(strProjet)

        Next

        If dicoProjet.ContainsKey(txt_Projet.Text) Then

            For Each strTache In dicoProjet(txt_Projet.Text).Keys

                txt_Tache.Items.Add(strTache)

            Next

        Else

            'txt_Tache.Text = ""
            'txt_SSTache.Text = ""
            txt_Tache.Items.Clear()
            txt_SSTache.Items.Clear()
            Exit Sub

        End If

        If dicoProjet(txt_Projet.Text).ContainsKey(txt_Tache.Text) Then

            For Each strSSTache In dicoProjet(txt_Projet.Text)(txt_Tache.Text).Keys

                txt_SSTache.Items.Add(strSSTache)

            Next

        Else

            'txt_SSTache.Text = ""
            txt_SSTache.Items.Clear()
            Exit Sub

        End If



    End Sub

End Class

Public Class Pointage

    Public Const fileCsvPath As String = "C:\Users\menan\source\repos\AutoPoint\AutoPoint.csv"

    Private csvEncoding As System.Text.Encoding = System.Text.Encoding.Unicode

    Private pStatus As String
    Private pDate As String
    Private pOpen As String
    Private pClose As String
    Private pLaps As String

    Private mProjet As String
    Private mTache As String
    Private mSSTache As String

    Public Sub Read(strLineInput As String)

        Dim ptDetail() As String

        ptDetail = Split(strLineInput, ";")

        If UBound(ptDetail) = 4 Then
            pStatus = "Open"
        ElseIf UBound(ptDetail) = 5 Then
            pStatus = "Closed"
        Else
            pStatus = "Error"
            Err.Raise(9999,, "Mauvais format de ligne en entrée : " & strLineInput)
            Exit Sub
        End If

        mProjet = ptDetail(0)
        mTache = ptDetail(1)
        mSSTache = ptDetail(2)
        pDate = ptDetail(3)
        pOpen = ptDetail(4)

        If UBound(ptDetail) > 4 Then
            pClose = ptDetail(5)
            pLaps = GetTimeLaps(Convert.ToDateTime(pClose))
        End If

    End Sub

    Public Sub Open(projet_ As String, tache_ As String, ssTache_ As String)

        pStatus = "Open"
        pDate = Format(Now, "dd/MM/yyyy")
        pOpen = Format(Now, "HH:mm")
        pClose = ""
        pLaps = "00:00:00"

        mProjet = projet_
        mTache = tache_
        mSSTache = ssTache_


        Dim mStreamWriter As StreamWriter
        mStreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fileCsvPath, True, csvEncoding)

        mStreamWriter.Write(mProjet & ";" & mTache & ";" & mSSTache & ";" & pDate & ";" & pOpen)

        mStreamWriter.Close()

    End Sub

    Public Sub Close()

        pStatus = "Closed"
        pClose = Format(Now, "HH:mm")
        pLaps = GetTimeLaps(Convert.ToDateTime(pClose))

        Dim mStreamWriter As StreamWriter
        mStreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fileCsvPath, True, csvEncoding)

        mStreamWriter.Write(";" & pClose & vbCrLf)

        mStreamWriter.Close()

    End Sub

    Public Function GetTimeLaps(closeTimer As DateTime) As String

        Dim openDate As DateTime = Convert.ToDateTime(pOpen)
        Dim lapsDate As TimeSpan = closeTimer - openDate

        Return Format(lapsDate.Hours, "00:") & Format(lapsDate.Minutes, "00")

    End Function

    Property strLog() As String
        Get
            Dim logTable() As String = Nothing

            Select Case pStatus

                Case "Open"

                    ReDim logTable(0 To 3)

                    logTable(0) = "Status : " & pStatus
                    logTable(1) = "Date : " & pDate
                    logTable(2) = "Open : " & pOpen
                    logTable(3) = "Temps : " & pLaps

                Case "Closed"

                    ReDim logTable(0 To 4)

                    logTable(0) = "Status : " & pStatus
                    logTable(1) = "Date : " & pDate
                    logTable(2) = "Open : " & pOpen
                    logTable(3) = "Close : " & pClose
                    logTable(4) = "Temps : " & pLaps

                Case "Error"

                    ReDim logTable(0)

                    logTable(0) = "Status : " & pStatus

            End Select



            Return Join(logTable, vbCrLf)
        End Get

        Set(value As String) : End Set
    End Property

    Property Status() As String
        Get
            Return pStatus
        End Get

        Set(value As String) : End Set
    End Property

    Property PointDate() As String
        Get
            Return pDate
        End Get

        Set(value As String) : End Set
    End Property

    Property Open_Hour() As String
        Get
            Return pOpen
        End Get

        Set(value As String) : End Set
    End Property

    Property Close_Hour() As String
        Get
            Return pClose
        End Get

        Set(value As String) : End Set
    End Property

    Property Duration() As String
        Get
            Return pLaps
        End Get

        Set(value As String)
            pLaps = value
        End Set
    End Property

    Property Projet() As String
        Get
            Return mProjet
        End Get

        Set(value As String) : End Set
    End Property
    Property Tache() As String
        Get
            Return mTache
        End Get

        Set(value As String) : End Set
    End Property

    Property SousTache() As String
        Get
            Return mSSTache
        End Get

        Set(value As String) : End Set
    End Property

End Class
