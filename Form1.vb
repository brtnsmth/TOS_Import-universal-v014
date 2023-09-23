Option Explicit On

' **********
' Adaptive data save mode select - recognizes system host


' TOS - Excel interface
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

' StreanWriter event log interface
Imports System
Imports System.IO
Imports System.Text

' SQL db interface
Imports System.Data
Imports System.Data.SqlClient

' pgr version interface
Imports System.Deployment
Imports System.Deployment.Application

Public Class Form1

#Region "TOS Dims"
    ' Define Excel source
    ReadOnly oXL As New Excel.Application
    Private oWB As Excel.Workbook
    Private oSheet As Excel.Worksheet
    Private oRng As Excel.Range

    ' Define Excel source cells
    Private XL_Cell_1, XL_Cell_2, XL_Cell_3, XL_Cell_4, XL_Cell_5 As Excel.Range
    Private XL_Cell_6, XL_Cell_7, XL_Cell_8, XL_Cell_9, XL_Cell_10 As Excel.Range
    Private XL_Cell_11, XL_Cell_12, XL_Cell_13, XL_Cell_14, XL_Cell_15 As Excel.Range
    Private XL_Cell_16, XL_Cell_17, XL_Cell_18, XL_Cell_19, XL_Cell_20 As Excel.Range
    Private XL_Cell_21, XL_Cell_22, XL_Cell_23 As Excel.Range
    Private XL_Cell_24, XL_Cell_25, XL_Cell_26 As Excel.Range
    Private XL_Cell_27, XL_Cell_28, XL_Cell_29, XL_Cell_30 As Excel.Range

    ' Define Excel source ticker
    Private Ticker_1, Ticker_2, Ticker_3, Ticker_4, Ticker_5 As String
    Private Ticker_6, Ticker_7, Ticker_8, Ticker_9, Ticker_10 As String
    Private Ticker_11, Ticker_12, Ticker_13, Ticker_14, Ticker_15 As String
    Private Ticker_16, Ticker_17, Ticker_18, Ticker_19, Ticker_20 As String
    Private Ticker_21, Ticker_22, Ticker_23 As String
    Private Ticker_24, Ticker_25, Ticker_26 As String
    Private Ticker_27, Ticker_28, Ticker_29, Ticker_30 As String

    ' Define DB output columns
    Private DB_Col_1, DB_Col_2, DB_Col_3, DB_Col_4, DB_Col_5 As String
    Private DB_Col_6, DB_Col_7, DB_Col_8, DB_Col_9, DB_Col_10 As String
    Private DB_Col_11, DB_Col_12, DB_Col_13, DB_Col_14, DB_Col_15 As String
    Private DB_Col_16, DB_Col_17, DB_Col_18, DB_Col_19, DB_Col_20 As String
    Private DB_Col_21, DB_Col_22, DB_Col_23 As String
    Private DB_Col_24, DB_Col_25, DB_Col_26 As String
    Private DB_Col_27, DB_Col_28, DB_Col_29, DB_Col_30 As String

    Private Sub Label41_Click(sender As Object, e As EventArgs) Handles Label41.Click

    End Sub

    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click

    End Sub

    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles lblValue_25.Click

    End Sub

    Private intOldVolume_ES As Integer = 0
    Private intNewVolume_ES As Integer = 0

    Private intOldVolume_TLT As Integer = 0
    Private intNewVolume_TLT As Integer = 0

    Private tsTimeStamp As DateTime
    Private intTimeToggle As Int16

    ' start mode variables
    Private boolDataCollStarted As Boolean = False
    Private strDayTime As String                    ' current time as hour minute to test against trigger
    Private strDayTimeTrigger As String = "1759"    ' look for minute before 6 PM to start new data file

    ' data coll of SMA 
    Private boolCollectSMA As Boolean = False       ' toggle for collection of SMA info
    Private decOldSMA20 As Decimal = 0              ' place holder for SMA data fail
    Private decOldSMA50 As Decimal = 0              ' place holder for SMA data fail

#End Region

#Region "SQL dims"

    '******************
    ' Phred5 SQL interface to database
    Dim cnnTOS_Phred5 As New SqlClient.SqlConnection With {.ConnectionString = "Data Source=Phred5\sqlexpress; Initial Catalog=TOS_Import; Integrated Security=True"}

    '******************
    ' Phred8 SQL interface to database
    Dim cnnTOS_Phred8 As New SqlConnection With {.ConnectionString = "Data Source=Phred8\SQL_Dev; Initial Catalog=TOS_Import; Integrated Security=True"}

    '******************
    ' PHRED9 SQL interface to database
    Dim cnnTOS_PHRED9 As New SqlClient.SqlConnection With {.ConnectionString = "Data Source=PHRED9\SQLEXPRESS_2019; Initial Catalog=TOS_Import; Integrated Security=True"}

    '******************
    ' PHRED11 SQL interface to database
    Dim cnnTOS_PHRED11 As New SqlClient.SqlConnection With {.ConnectionString = "Data Source=PHRED11\SQLEXPRESS; Initial Catalog=TOS_Import; Integrated Security=True"}


    '******************
    Private cmdTOS_RPD_DataINSERT As New SqlClient.SqlCommand
    Private cmd_TOS_TickersINSERT As New SqlClient.SqlCommand

#End Region

#Region "CSV save dims"

    '******************
    ' StreamWriter interface to EventLog file
    Private swCSV As StreamWriter

    Private strCSV_FileNameBase As String = "C:\temp\TOS import CSV files\" & strPC_Name & " CSV data "      ' root of filename for streamwriter
    Private strCSV_FileName As String                                                     ' section of filename built JIT
    Private strCSV_FileNameExtension As String = ".CSV"                                   ' defines file as CSV(txt) format
    Private strTimeStamp As String                  ' time string used to create unique data file
    Private strCSV_Data As String                   ' data to be logged
    Private boolNewCSVFileCreated As Boolean = False   ' check to see if file exists for appending or if need to create
    Private strCSV_Header As String = "Date" & ", " & "/ES" & ", " & "/NQ" & ", " & "/RTY" & ", " & "SPY" & ", " & "QQQ" & ", " & "IWM" & ", " & "AAPL" & ", " & "MSFT" & ", " & "NVDA" & ", " & "XLK" & ", " & "XLF" & ", " & "XLP" & ", " & "XLY" & ", " & "XTN" & ", " & "HYG" & ", " & "TNX" & ", " & "TYX" & ", " & "/ES volume" & ", " & "TLT" & ", " & "TLT volume" & ", " & "VIX" & ", " & "SPX" & ", " & "SPX PCR" & ", " & "SPY PCR" & ", " & "/ES PCR"

#End Region

#Region "Event log dims"
    ' *****************
    ' system info
    Private strPC_Name As String = System.Windows.Forms.SystemInformation.ComputerName
    Private strPgrVersionNumber As String

    ' watchdog timer values for comparison
    Private decOldValue_ES As Decimal
    Private decNewValue_ES As Decimal
    Private boolWD_Warning As Boolean = False
    Private boolWD_WarningMsgSent As Boolean = False
    Private boolWD_OK As Boolean = False
    Private boolWD_OKMsgSent As Boolean = False

    ' tick-tock colors
    Private clrTick As Color = Color.LightGray
    Private clrTock As Color = Color.SlateGray
    Private clrNormalLight As Color = Color.LightGray
    Private clrNormalDark As Color = Color.SlateGray
    Private clrWarning As Color = Color.Lime


    '******************
    ' StreamWriter interface to EventLog file
    Private swLog As StreamWriter

    Private strLog_FileNameBase As String = "C:\temp\TOS import log files\"               ' root of filename for streamwriter
    Private strLog_FileName As String                                                     ' section of filename built JIT
    Private strLog_FileNameExtension As String = ".log"                                   ' defines file as log(txt) format
    Private strEventInfo As String                  ' event info to be logged
    Private boolNewLogFileCreated As Boolean = False   ' check to see if file exists for appending or if need to create
    Private strSaveMode As String

#End Region

#Region "Form Load / Close"


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Initialize screen labels
        Call PgrVersionNumber()
        Call InitializeScreen()

    End Sub

    Private Sub PgrVersionNumber()

        If (ApplicationDeployment.IsNetworkDeployed) Then
            Dim AD As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
            strPgrVersionNumber = strPC_Name & " - " & strSaveMode & " - pgr version " & AD.CurrentVersion.ToString
        Else
            strPgrVersionNumber = strPC_Name & " - " & strSaveMode & " - pgr version " & My.Application.Info.Version.ToString
        End If

        lblPgrVersionNumber.Text = strPgrVersionNumber


    End Sub


    Private Sub btnStartModeSelected_Click(sender As Object, e As EventArgs) Handles btnStartModeSelected.Click

        btnStartModeSelected.Enabled = False

        If rbSaveAs_CSV.Checked Then
            strSaveMode = "CSV"
        ElseIf rbSaveAs_SQL.Checked Then
            strSaveMode = "SQL"
        End If

        '****************************************

        ' start StreamWriter
        Call CreateNewLogFile()

        ' Set data coll start mode
        If rbImmediateStart.Checked Then

            '******************
            strEventInfo = "Immediate pgr start enabled."
            Call LogEvent()

            '******************
            Call DataCollStart()
        Else

            '******************
            strEventInfo = "Time trigger pgr start enabled."
            Call LogEvent()

            '******************
            ' enable time check
            tmrTimeCheck.Enabled = True
            lblStartStatus.Text = "Waiting for start time."
            lblTargetTime.Text = strDayTimeTrigger
        End If

    End Sub

    Private Sub tmrTimeCheck_Tick(sender As Object, e As EventArgs) Handles tmrTimeCheck.Tick

        ' reset timer
        tmrTimeCheck.Enabled = False
        tmrTimeCheck.Enabled = True

        ' get current time
        strDayTime = Format(Now(), "HH" & "mm")
        lblCurrentTime.Text = Format(Now(), "HH" & "." & "mm" & "." & "ss")

        'test against time trigger
        If strDayTime = strDayTimeTrigger And boolDataCollStarted = False Then
            Call DataCollStart()
            boolDataCollStarted = True

        End If

        If strDayTime <> strDayTimeTrigger And boolDataCollStarted = True Then
            boolDataCollStarted = False
            tmrTimeCheck.Enabled = False
        End If

        Debug.Print("boolDataCollStarted = " & boolDataCollStarted)

    End Sub




    Private Sub DataCollStart()

        ' init messaging
        pnlTickTock.Visible = True
        lblTickTock.Text = "Starting XL"
        lblStartStatus.Text = "Starting data coll."


        ' ***************************************
        ' Initialize Excel workbook
        oXL.Visible = False
        oWB = oXL.Workbooks.Add
        oSheet = oWB.ActiveSheet

        ' run delay so Excel  can load 
        tmrExcelLoad.Interval = 5000
        tmrExcelLoad.Start()
        If tmrExcelLoad.Enabled = True Then Debug.Print("Excel Load wait timer started")

        ' ***************************************
        ' Load variables and Excel formulas

        '  Call Load_Tickers()
        Call Define_Cell_Ranges()
        Call Define_Cell_Formulas()

        '****************************************
        ' check data save method is CSV (SQL needs no new file)
        If rbSaveAs_CSV.Checked Then
            ' start CSV StreamWriter
            Call CreateNewCSVFile()
        End If


        lblStartStatus.Text = "Data coll running."

    End Sub




    Private Sub InitializeScreen()

        ' Initialize labels

        lblTickTock.Text = ""
        intTimeToggle = 0
        lblDataWrite_Status.Text = ""
        lblDataWrite_TimeStamp.Text = ""
        lblStartStatus.Text = "Waiting for selection."
        lblTargetTime.Text = ""
        lblCurrentTime.Text = ""
        rbTimeTriggerStart.Checked = True
        rbSaveAs_SQL.Checked = True
        pnlTickTock.Visible = False


        'If strSaveMode = "CSV" Then
        '    Me.Text = "CSV saves form"
        'ElseIf strSaveMode = "SQL" Then
        '    Me.Text = "SQL saves form"
        'End If

        ' Clear values returned locally
        lblValue_1.Text = ""
        lblValue_2.Text = ""
        lblValue_3.Text = ""
        lblValue_4.Text = ""
        lblValue_5.Text = ""
        lblValue_6.Text = ""
        lblValue_7.Text = ""
        lblValue_8.Text = ""
        lblValue_9.Text = ""
        lblValue_10.Text = ""

        lblValue_11.Text = ""
        lblValue_12.Text = ""
        lblValue_13.Text = ""
        lblValue_14.Text = ""
        lblValue_15.Text = ""
        lblValue_16.Text = ""
        lblValue_17.Text = ""
        lblValue_18.Text = ""
        lblValue_19.Text = ""
        lblValue_20.Text = ""

        lblValue_21.Text = ""
        lblValue_22.Text = ""
        lblValue_23.Text = ""
        lblValue_24.Text = ""
        lblValue_25.Text = ""

    End Sub


    Private Sub Define_Cell_Ranges()

        ' define cell ranges
        XL_Cell_1 = oSheet.Cells(1, 1)
        XL_Cell_2 = oSheet.Cells(2, 1)
        XL_Cell_3 = oSheet.Cells(3, 1)
        XL_Cell_4 = oSheet.Cells(4, 1)
        XL_Cell_5 = oSheet.Cells(5, 1)
        XL_Cell_6 = oSheet.Cells(6, 1)
        XL_Cell_7 = oSheet.Cells(7, 1)
        XL_Cell_8 = oSheet.Cells(8, 1)
        XL_Cell_9 = oSheet.Cells(9, 1)
        XL_Cell_10 = oSheet.Cells(10, 1)

        XL_Cell_11 = oSheet.Cells(11, 1)
        XL_Cell_12 = oSheet.Cells(12, 1)
        XL_Cell_13 = oSheet.Cells(13, 1)
        XL_Cell_14 = oSheet.Cells(14, 1)
        XL_Cell_15 = oSheet.Cells(15, 1)
        XL_Cell_16 = oSheet.Cells(16, 1)
        XL_Cell_17 = oSheet.Cells(17, 1)
        XL_Cell_18 = oSheet.Cells(18, 1)
        XL_Cell_19 = oSheet.Cells(19, 1)
        XL_Cell_20 = oSheet.Cells(20, 1)

        XL_Cell_21 = oSheet.Cells(21, 1)
        XL_Cell_22 = oSheet.Cells(22, 1)
        XL_Cell_23 = oSheet.Cells(23, 1)
        XL_Cell_24 = oSheet.Cells(24, 1)
        XL_Cell_25 = oSheet.Cells(25, 1)
        XL_Cell_26 = oSheet.Cells(26, 1)
        XL_Cell_27 = oSheet.Cells(27, 1)
        XL_Cell_28 = oSheet.Cells(28, 1)
        XL_Cell_29 = oSheet.Cells(29, 1)
        XL_Cell_30 = oSheet.Cells(30, 1)

    End Sub


    Private Sub Define_Cell_Formulas()

        Try
            ' define cell formulas
            XL_Cell_1.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/ES:XCME"" )"
            XL_Cell_2.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/NQ:XCME"" )"
            XL_Cell_3.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/RTY:XCME"" )"
            XL_Cell_4.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""SPY"" )"
            XL_Cell_5.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""QQQ"" )"
            XL_Cell_6.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""IWM"" )"
            XL_Cell_7.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""AAPL"" )"
            XL_Cell_8.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""MSFT"" )"
            XL_Cell_9.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""NVDA"" )"
            XL_Cell_10.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLK"" )"

            XL_Cell_11.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLF"" )"
            XL_Cell_12.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLP"" )"
            XL_Cell_13.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XLY"" )"
            XL_Cell_14.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""XTN"" )"
            XL_Cell_15.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""HYG"" )"

            XL_Cell_16.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""TNX"" )"
            XL_Cell_17.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""TYX"" )"

            XL_Cell_18.Formula = "=RTD(""TOS.RTD"", , ""VOLUME"", ""/ES:XCME"" )"
            XL_Cell_19.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""TLT"" )"
            XL_Cell_20.Formula = "=RTD(""TOS.RTD"", , ""VOLUME"", ""TLT"" )"

            XL_Cell_21.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""VIX"" )"
            XL_Cell_22.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""SPX"" )"
            XL_Cell_23.Formula = "=RTD(""TOS.RTD"", , ""Put_Call_Ratio"", ""SPX"" )"
            XL_Cell_24.Formula = "=RTD(""TOS.RTD"", , ""Put_Call_Ratio"", ""SPY"" )"
            XL_Cell_25.Formula = "=RTD(""TOS.RTD"", , ""Put_Call_Ratio"", ""/ES:XCME"" )"

            Debug.Print("RTD formulas loaded OK")

            '******************
            strEventInfo = "RTD formulas loaded OK."
            Call LogEvent()

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

            ' System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL)
        End Try


    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        strEventInfo = "Form closing."
        Call LogEvent()

        oRng = Nothing
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL)

        ' TODO -find way to close Excel without external Save? msgbox
    End Sub

#End Region

#Region "Timer events"

    Private Sub tmrExcelLoad_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrExcelLoad.Tick

        '****************************************
        ' Excel load delay before connecting to TOS RTD is done 
        Debug.Print("Excel Load wait timer done")

        '******************
        strEventInfo = "Excel Load wait timer done."
        Call LogEvent()

        tmrExcelLoad.Enabled = False

        '****************************************
        ' Test VB - Excel RTD connection
        Call Test_Connection()

        '****************************************
        ' VB - Excel RTD connection OK
        ' start main looping
        tmrMainLoop.Interval = 3 * intSecond
        tmrMainLoop.Start()


        Debug.Print("Main data collection loop active.")

        '******************
        strEventInfo = "Main data collection loop active."
        Call LogEvent()

        '****************************************
        ' start main loop watchdog timer
        tmrWatchdog.Interval = 15 * intSecond
        tmrWatchdog.Enabled = True

        Debug.Print("Main loop watchdog timer active.")

        '******************
        strEventInfo = "Main loop watchdog timer active."
        Call LogEvent()

    End Sub

    Private Sub Test_Connection()

        Try
            Debug.Print("in Test_Connection")
            Dim rg As Excel.Range = oSheet.Cells(1, 1)

            rg.Formula = "=RTD(""TOS.RTD"", , ""LAST"", ""/ES:XCME"" )"

            If CStr(rg.Value) = "" Then
                Beep()
                MessageBox.Show("No connection to TOS RTD server." & vbCrLf & "Exit pgr, start TOS, restart pgr.")

                '******************
                strEventInfo = "No connection to TOS RTD server. > Exit pgr, start TOS, restart pgr."
                Call LogEvent()
            Else
                Debug.Print("RTD connection tested OK")

                '******************
                strEventInfo = "RTD connection tested OK."
                Call LogEvent()

                lblTickTock.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.Message)

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

            ' System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL)

        End Try


    End Sub

    Private Sub tmrMainLoop_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMainLoop.Tick

        tmrMainLoop.Enabled = False
        tmrMainLoop.Enabled = True

        If intTimeToggle = 0 Then
            pnlTickTock.BackColor = clrTick
            lblTickTock.Text = "      tick"
            intTimeToggle = 1
        Else
            pnlTickTock.BackColor = clrTock
            lblTickTock.Text = "      tock"
            intTimeToggle = 0
        End If

        Call Get_Values()
        Call Display_Values()

        '****************************************
        ' check data save method
        If rbSaveAs_CSV.Checked Then
            Call CSV_Save()
        ElseIf rbSaveAs_SQL.Checked Then
            Call DB_Save()
        End If

        lblCurrentTime.Text = Format(Now(), "HH" & "." & "mm" & "." & "ss")

    End Sub


    Private Sub tmrWatchdog_Tick(sender As Object, e As EventArgs) Handles tmrWatchdog.Tick

        Try
            ' check for value change in /ES volume
            If CStr(XL_Cell_18.Value) = "N/A" Then
                decNewValue_ES = 0
            Else
                decNewValue_ES = CDec(XL_Cell_18.Value)
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
            Debug.Print("exit exception at " & Now())

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

        End Try

        If decNewValue_ES = decOldValue_ES Then
            ' indicates data stream issue
            clrTock = clrWarning            ' Lime
            boolWD_Warning = True
            boolWD_OK = False
        Else
            clrTock = clrNormalDark         ' SlateGray
            boolWD_Warning = False
            boolWD_OK = True
        End If

        ' Remember this value for next interval check
        decOldValue_ES = decNewValue_ES

        ' log events
        If boolWD_OK = True And boolWD_OKMsgSent = False Then
            strEventInfo = "Main loop watchdog timer OK."
            Call LogEvent()
            boolWD_OKMsgSent = True
            boolWD_WarningMsgSent = False
            tmrNotification.Enabled = False
        ElseIf boolWD_Warning = True And boolWD_WarningMsgSent = False Then
            strEventInfo = "Main loop watchdog timer WARNING active."
            Call LogEvent()
            boolWD_OKMsgSent = False
            boolWD_WarningMsgSent = True
            ' set audible notification
            Beep()
            tmrNotification.Interval = 15000
            tmrNotification.Enabled = True
        End If



    End Sub


    Private Sub btnSilenceWarning_Click(sender As Object, e As EventArgs) Handles btnSilenceWarning.Click

        tmrSilenceWarning.Interval = 900 * intSecond
        tmrSilenceWarning.Enabled = True
        btnSilenceWarning.BackColor = Color.Lime
        btnSilenceWarning.Text = "Warning" & vbCrLf & "Silenced"

    End Sub

    Private Sub tmrSilenceWarning_Tick(sender As Object, e As EventArgs) Handles tmrSilenceWarning.Tick

        tmrSilenceWarning.Enabled = False
        btnSilenceWarning.BackColor = Color.Gray
        btnSilenceWarning.Text = "Silence" & vbCrLf & "Warning"

    End Sub

    Private Sub tmrNotification_Tick(sender As Object, e As EventArgs) Handles tmrNotification.Tick

        If tmrSilenceWarning.Enabled Then
            ' don't beep
        Else
            Beep()
        End If

    End Sub


#End Region

#Region "display data"

    Private Sub Get_Values()

        Try

            ' Load values returned from RTDServer
            DB_Col_1 = CStr(XL_Cell_1.Value)
            DB_Col_2 = CStr(XL_Cell_2.Value)
            DB_Col_3 = CStr(XL_Cell_3.Value)
            DB_Col_4 = CStr(XL_Cell_4.Value)
            DB_Col_5 = CStr(XL_Cell_5.Value)
            DB_Col_6 = CStr(XL_Cell_6.Value)
            DB_Col_7 = CStr(XL_Cell_7.Value)
            DB_Col_8 = CStr(XL_Cell_8.Value)
            DB_Col_9 = CStr(XL_Cell_9.Value)
            DB_Col_10 = CStr(XL_Cell_10.Value)

            DB_Col_11 = CStr(XL_Cell_11.Value)
            DB_Col_12 = CStr(XL_Cell_12.Value)
            DB_Col_13 = CStr(XL_Cell_13.Value)
            DB_Col_14 = CStr(XL_Cell_14.Value)
            DB_Col_15 = CStr(XL_Cell_15.Value)

            DB_Col_16 = CStr(XL_Cell_16.Value)
            DB_Col_17 = CStr(XL_Cell_17.Value)

            ' calc interval volume for /ES
            If CStr(XL_Cell_18.Value) = "N/A" Then
                DB_Col_18 = 0
            Else
                intNewVolume_ES = CInt(XL_Cell_18.Value)
                DB_Col_18 = intNewVolume_ES - intOldVolume_ES
                intOldVolume_ES = intNewVolume_ES
            End If

            DB_Col_19 = CStr(XL_Cell_19.Value)

            ' calc interval volume for TLT
            If CStr(XL_Cell_20.Value) = "N/A" Then
                DB_Col_20 = 0
            Else
                intNewVolume_TLT = CInt(XL_Cell_20.Value)
                DB_Col_20 = intNewVolume_TLT - intOldVolume_TLT
                intOldVolume_TLT = intNewVolume_TLT

            End If

            DB_Col_21 = CStr(XL_Cell_21.Value)
            DB_Col_22 = CStr(XL_Cell_22.Value)

            If IsNumeric(XL_Cell_23.Value) Then
                DB_Col_23 = CStr(XL_Cell_23.Value)
            Else
                DB_Col_23 = "0"
            End If

            If IsNumeric(XL_Cell_24.Value) Then
                DB_Col_24 = CStr(XL_Cell_24.Value)
            Else
                DB_Col_24 = "0"
            End If

            If IsNumeric(XL_Cell_25.Value) Then
                DB_Col_25 = CStr(XL_Cell_25.Value)
            Else
                DB_Col_25 = "0"
            End If

        Catch ex As Exception
            Debug.Print(ex.Message)
            Debug.Print("exit exception at " & Now())

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

        End Try



    End Sub

    Private Sub Display_Values()


        ' Display values returned locally
        lblValue_1.Text = DB_Col_1
        lblValue_2.Text = DB_Col_2
        lblValue_3.Text = DB_Col_3
        lblValue_4.Text = DB_Col_4
        lblValue_5.Text = DB_Col_5
        lblValue_6.Text = DB_Col_6
        lblValue_7.Text = DB_Col_7
        lblValue_8.Text = DB_Col_8
        lblValue_9.Text = DB_Col_9
        lblValue_10.Text = DB_Col_10

        lblValue_11.Text = DB_Col_11
        lblValue_12.Text = DB_Col_12
        lblValue_13.Text = DB_Col_13
        lblValue_14.Text = DB_Col_14
        lblValue_15.Text = DB_Col_15
        lblValue_16.Text = DB_Col_16
        lblValue_17.Text = DB_Col_17
        lblValue_18.Text = DB_Col_18
        lblValue_19.Text = DB_Col_19
        lblValue_20.Text = DB_Col_20

        lblValue_21.Text = DB_Col_21
        lblValue_22.Text = DB_Col_22
        lblValue_23.Text = DB_Col_23
        lblValue_24.Text = DB_Col_24
        lblValue_25.Text = DB_Col_25

    End Sub


#End Region


#Region "DB Saves"

    Private Sub DB_Save()


        tsTimeStamp = Now()


        '  **********************
        '   open connection to SQL database           
        Try
            If strPC_Name = "PHRED9" Then
                cnnTOS_PHRED9.Open()
            ElseIf strPC_Name = "PHRED8" Then
                cnnTOS_Phred8.Open()
            ElseIf strPC_Name = "PHRED5" Then
                cnnTOS_Phred5.Open()
            ElseIf strPC_Name = "PHRED11" Then
                cnnTOS_PHRED11.Open()
            End If


            cmdTOS_RPD_DataINSERT.CommandText = "INSERT TOS_ImportData (TimeStamp, DB_Col_1, DB_Col_2, DB_Col_3, DB_Col_4, DB_Col_5, DB_Col_6, DB_Col_7, DB_Col_8, DB_Col_9, DB_Col_10, DB_Col_11, DB_Col_12, DB_Col_13, DB_Col_14, DB_Col_15, DB_Col_16, DB_Col_17, DB_Col_18, DB_Col_19, DB_Col_20, DB_Col_21, DB_Col_22, DB_Col_23, DB_Col_24, DB_Col_25)" _
               & "VALUES( '" & tsTimeStamp & "', '" & DB_Col_1 & "', '" & DB_Col_2 & "', '" & DB_Col_3 & "', '" & DB_Col_4 & "', '" & DB_Col_5 & "', '" & DB_Col_6 & "', '" & DB_Col_7 & "', '" & DB_Col_8 & "', '" & DB_Col_9 & "', '" & DB_Col_10 & "', '" & DB_Col_11 & "', '" & DB_Col_12 & "', '" & DB_Col_13 & "', '" & DB_Col_14 & "', '" & DB_Col_15 & "', '" & DB_Col_16 & "', '" & DB_Col_17 & "', '" & DB_Col_18 & "', '" & DB_Col_19 & "', '" & DB_Col_20 & "', '" & DB_Col_21 & "', '" & DB_Col_22 & "', '" & DB_Col_23 & "', '" & DB_Col_24 & "', '" & DB_Col_25 & "' )"

            If strPC_Name = "PHRED9" Then
                cmdTOS_RPD_DataINSERT.Connection = cnnTOS_PHRED9
            ElseIf strPC_Name = "PHRED8" Then
                cmdTOS_RPD_DataINSERT.Connection = cnnTOS_Phred8
            ElseIf strPC_Name = "PHRED5" Then
                cmdTOS_RPD_DataINSERT.Connection = cnnTOS_Phred5
            ElseIf strPC_Name = "PHRED11" Then
                cmdTOS_RPD_DataINSERT.Connection = cnnTOS_PHRED11
            End If

            cmdTOS_RPD_DataINSERT.ExecuteNonQuery()

            lblDataWrite_Status.Text = "SQL data INSERT success."
            lblDataWrite_TimeStamp.Text = Now()

        Catch ex As Exception

            tmrMainLoop.Enabled = False

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

            ' System.Windows.Forms.MessageBox.Show(ex.Message)
            Debug.Print("INSERT failed")

            lblDataWrite_Status.Text = "SQL data INSERT failed."
            lblDataWrite_TimeStamp.Text = Now()

            If strPC_Name = "PHRED9" Then
                cnnTOS_PHRED9.Close()
            ElseIf strPC_Name = "PHRED8" Then
                cnnTOS_Phred8.Close()
            ElseIf strPC_Name = "PHRED5" Then
                cnnTOS_Phred5.Close()
            ElseIf strPC_Name = "PHRED11" Then
                cnnTOS_PHRED11.Close()
            End If

        Finally

            If strPC_Name = "PHRED9" Then
                cnnTOS_PHRED9.Close()
            ElseIf strPC_Name = "PHRED8" Then
                cnnTOS_Phred8.Close()
            ElseIf strPC_Name = "PHRED5" Then
                cnnTOS_Phred5.Close()
            ElseIf strPC_Name = "PHRED11" Then
                cnnTOS_PHRED11.Close()
            End If

        End Try


    End Sub


#End Region


#Region "CSV saves"

    Private Sub CreateNewCSVFile()

        '****************************************
        ' build timestamp to embed in filename
        strTimeStamp = Format(Now(), "yyyy" & "-" & "MM" & "-" & "dd" & " " & "HH" & "." & "mm" & "." & "ss")
        Try
            ' build filename
            strCSV_FileName = strCSV_FileNameBase & strTimeStamp & strCSV_FileNameExtension
            ' Open NEW StreamWriter
            swCSV = My.Computer.FileSystem.OpenTextFileWriter(strCSV_FileName, True)
            ' write column headers to first record
            swCSV.WriteLine("Instance of CSV data save opened " & strTimeStamp)
            swCSV.WriteLine(strPgrVersionNumber)
            swCSV.WriteLine()
            swCSV.WriteLine(strCSV_Header)
            swCSV.Close()
        Catch ex As IOException

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

            ' System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try


        boolNewCSVFileCreated = True

    End Sub


    Private Sub CSV_Save()

        ' get current time
        tsTimeStamp = Now()

        '****************************************
        ' build data string
        strCSV_Data = CStr(tsTimeStamp) & ", " & DB_Col_1 & ", " & DB_Col_2 & ", " & DB_Col_3 & ", " & DB_Col_4 & ", " & DB_Col_5 & ", " & DB_Col_6 & ", " & DB_Col_7 & ", " & DB_Col_8 & ", " & DB_Col_9 & ", " & DB_Col_10 & ", " & DB_Col_11 & ", " & DB_Col_12 & ", " & DB_Col_13 & ", " & DB_Col_14 & ", " & DB_Col_15 & ", " & DB_Col_16 & ", " & DB_Col_17 & ", " & DB_Col_18 & ", " & DB_Col_19 & ", " & DB_Col_20 & ", " & DB_Col_21 & ", " & DB_Col_22 & ", " & DB_Col_23 & ", " & DB_Col_24 & ", " & DB_Col_25

        ' write data 

        '****************************************
        ' Open NEW StreamWriter and write column headers to first line.

        Try
            swCSV = My.Computer.FileSystem.OpenTextFileWriter(strCSV_FileName, True)
            swCSV.WriteLine(strCSV_Data)
            swCSV.Close()

            ' Debug.Print("CSV data write success @ " & Now())
            lblDataWrite_Status.Text = "CSV data write success."
            lblDataWrite_TimeStamp.Text = Now()

        Catch ex As IOException

            '******************
            strEventInfo = ex.Message
            Call LogEvent()

            ' System.Windows.Forms.MessageBox.Show(ex.Message)
            lblDataWrite_Status.Text = "CSV data write failed."
            lblDataWrite_TimeStamp.Text = Now()
        End Try


    End Sub
#End Region


#Region "Event Logging"


    Private Sub CreateNewLogFile()

        '****************************************
        ' build filename
        strTimeStamp = Format(Now(), "yyyy" & "-" & "MM" & "-" & "dd" & " " & "HH" & "." & "mm" & "." & "ss")
        strLog_FileNameBase = strLog_FileNameBase & strPC_Name & " " & strSaveMode & " EventLog "

        Try
            ' build filename
            strLog_FileName = strLog_FileNameBase & strTimeStamp & strLog_FileNameExtension
            ' Open NEW StreamWriter
            swLog = My.Computer.FileSystem.OpenTextFileWriter(strLog_FileName, True)
            ' write header to first record
            swLog.WriteLine("Instance of event log opened " & strTimeStamp)
            swLog.WriteLine(strPgrVersionNumber)
            swLog.WriteLine()
            swLog.Close()
        Catch ex As IOException
            MsgBox(ex.ToString)
        End Try


        boolNewLogFileCreated = True

    End Sub

    Private Sub LogEvent()

        Dim strLog_Data As String

        '****************************************
        ' get current time
        tsTimeStamp = Now()
        ' build data string
        strLog_Data = CStr(tsTimeStamp) & " - " & strEventInfo


        '****************************************
        ' Open EventLog StreamWriter and write data
        If strLog_FileName <> "" Then       ' check for file initialized
            Try
                swLog = My.Computer.FileSystem.OpenTextFileWriter(strLog_FileName, True)
                swLog.WriteLine(strLog_Data)
                swLog.Close()

                lblEventLog_Timestamp.Text = CStr(tsTimeStamp)
                lblEventLog_Data.Text = strEventInfo

            Catch ex As IOException
                MsgBox(ex.ToString)

                ' System.Windows.Forms.MessageBox.Show(ex.Message)
                lblEventLog_Timestamp.Text = Now()
                lblEventLog_Data.Text = "Log data write failed @ "
            End Try
        End If

    End Sub


#End Region


End Class