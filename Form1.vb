Imports System.IO

Public Class Form1

    Private Const fileCsvPath As String = "C:\Users\menan\source\repos\AutoPoint\AutoPoint.csv"

    Private lastPointage As Pointage

    Public lastPoint As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim lastCsvLine As String

        lastPointage = New Pointage

        lastCsvLine = GetLastPoint()

        lastPointage.Read(lastCsvLine)

        Call WriteLog()

        Call Timer1_Tick()

    End Sub

    Private Sub btn_OpenPoint_Click(sender As Object, e As EventArgs) Handles btn_OpenPoint.Click

        If lastPointage Is Nothing Then Exit Sub

        Select Case lastPointage.Status

            Case "Closed"

                If txt_Tache.Text = "" Or txt_SSTache.Text = "" Then

                    MessageBox.Show("Veuillez renseigner une tâche + une sous-tâche.")
                    Exit Sub

                End If

                Timer1.Start()

                lastPointage = New Pointage
                lastPointage.Open(txt_Tache.Text, txt_SSTache.Text)

                txt_Tache.Enabled = False
                txt_SSTache.Enabled = False

            Case "Open"

                lastPointage.Close()
                txt_Tache.Enabled = True
                txt_SSTache.Enabled = True

        End Select

        Call WriteLog()

    End Sub

    Private Sub WriteLog()

        txt_Tache.Text = lastPointage.Tache
        txt_SSTache.Text = lastPointage.SousTache

        logFrame.Text = lastPointage.strLog

    End Sub

    Private Function GetLastPoint() As String

        Dim mStreamReader As StreamReader
        Dim tablePointFromCsv() As String

        mStreamReader = My.Computer.FileSystem.OpenTextFileReader(fileCsvPath)

        tablePointFromCsv = Split(mStreamReader.ReadToEnd, vbCrLf)

        mStreamReader.Close()

        If tablePointFromCsv(UBound(tablePointFromCsv)) = "" Then
            Return tablePointFromCsv(UBound(tablePointFromCsv) - 1)
        Else
            Return tablePointFromCsv(UBound(tablePointFromCsv))
        End If

    End Function

    Private Sub Timer1_Tick() Handles Timer1.Tick 'sender As Object, e As EventArgs

        With lastPointage

            If .Status = "Open" Then

                Dim openDate As DateTime = Convert.ToDateTime(.Open_Hour)
                Dim lapsDate As TimeSpan = Now - openDate

                .Duration = Format(lapsDate.Hours, "00:") & Format(lapsDate.Minutes, "00:") & Format(lapsDate.Seconds, "00")

                Me.Text = .Duration

                Call WriteLog()

            End If

        End With

    End Sub
End Class

Class Pointage

    Public Const fileCsvPath As String = "C:\Users\menan\source\repos\AutoPoint\AutoPoint.csv"

    Private pStatus As String
    Private pDate As String
    Private pOpen As String
    Private pClose As String
    Private pLaps As String

    Private mTache As String
    Private mSSTache As String

    Public Sub Read(strLineInput As String)

        Dim ptDetail() As String

        ptDetail = Split(strLineInput, ";")

        If UBound(ptDetail) = 3 Then
            pStatus = "Open"
        ElseIf UBound(ptDetail) = 4 Then
            pStatus = "Closed"
        Else
            pStatus = "Error"
            Err.Raise(9999,, "Mauvais format de ligne en entrée : " & strLineInput)
            Exit Sub
        End If

        mTache = ptDetail(0)
        mSSTache = ptDetail(1)
        pDate = ptDetail(2)
        pOpen = ptDetail(3)

        If UBound(ptDetail) > 3 Then
            pClose = ptDetail(4)
            pLaps = GetTimeLaps(Convert.ToDateTime(pClose))
        End If

    End Sub

    Public Sub Open(tache_ As String, ssTache_ As String)

        pStatus = "Open"
        pDate = Format(Now, "dd/MM/yyyy")
        pOpen = Format(Now, "HH:mm")
        pClose = ""
        pLaps = "00:00:00"
        mTache = tache_
        mSSTache = ssTache_


        Dim mStreamWriter As StreamWriter
        mStreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fileCsvPath, True)

        mStreamWriter.Write(mTache & ";" & mSSTache & ";" & pDate & ";" & pOpen)

        mStreamWriter.Close()

    End Sub

    Public Sub Close()

        Dim timeLaps(0 To 1) As Double

        pStatus = "Closed"
        pClose = Format(Now, "HH:mm")
        pLaps = GetTimeLaps(Convert.ToDateTime(pClose))

        Dim mStreamWriter As StreamWriter
        mStreamWriter = My.Computer.FileSystem.OpenTextFileWriter(fileCsvPath, True)

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
