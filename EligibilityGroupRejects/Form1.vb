Option Explicit On

Imports System.Deployment.Application
Imports System.IO
Imports System.Diagnostics
Imports System.Net.Mail
Imports bgw = System.ComponentModel
Imports pcom = AutPSTypeLibrary

Public Class Form1
    Dim objRx As pcom.AutPS
    Dim objWait As Object
    Dim objMgr, objMgr2 As Object
    Dim ObjSessionHandle As Integer
    Dim intSessions As Integer, x As Integer
    Dim autECLConnList As Object

    Dim usrNm As String

    'Excel Object Variables
    Dim objExcel
    Dim objWorkbook1
    Dim objWorksheet1
    Dim objWorksheet2
    Dim SSRowNumber As Integer
    Dim SSTabType As String  'G=Group and M=Member
    Dim JobName As String

    'Dim objExcelAppObject
    Dim objExcelfolder
    Dim objExcelDirectory
    Dim objExcelFilePath
    Dim msoFileDialogFolderPicker
    Dim var_SplitString
    Dim int_UBoundIndex
    Dim var_FileName As String

    'Script variables
    Dim int_Counter_Main       'This will walk us thru each RxNumber that needs to get Reversed and ReSubmitted
    Dim int_Counter_Summary    'This will walk us thru each row on the Summary page (each RXNumber will get 2 rows on the Summary tab)
    Dim rejRow, z

    Dim var_TimeStamp
    Dim objWMIService
    Dim objProcess
    Dim colProcess
    Dim str_Computer
    Dim str_ProcessToKill
    Dim int_ProcessRxClaimNbrListRowCount
    Dim int_ProcessRxClaimNbr

    Dim StartedAt As DateTime
    Dim FinishedAt As DateTime

    Dim int_TabNum As Integer   'This will keep track of which tab we are on in the spreadsheet

    ''*************  This will give us the Computer Name   ************************************************************************
    Dim objSysInfo = CreateObject("WinNTSystemInfo")
    Dim strComputerName = objSysInfo.ComputerName
    ''*****************************************************************************************************************************

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Dim iCarrierType As Integer

    Dim Carriers As New List(Of Carrier)
    Dim Clients As New List(Of Client)

    'I added this as a test for Sprint_2  1/27/16   (RLB)

    '*********************************************************************************************************************************************************************
    'Just added this to see if it shows up in GIT 1/27/16 (RLB)

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Default the date to today's date minus 1
        Me.DateTimePickerFrom.Value = Date.Today.AddDays(-1)

        'This will get the version we are on
        Try
            Dim ver As Version = ApplicationDeployment.CurrentDeployment.CurrentVersion
            Me.Text = "Eligibility Group Rejects   (vers " & ver.ToString & ")"
        Catch
            Me.Text = "MY Eligibility Group Rejects"
        End Try

        'This will open up our template and populate the checkbox list with valid clients
        Try
            OpenSpreadsheetTemplate()

            PopulateClientList()

            'default to all be checked
            Me.cbAll.Checked = True
        Catch
            Me.Text = "Failed opening template and popluating client checkbox list"
        End Try

        Try
            'this will make it so that we can use a background worker thread running (using) fields on the screen that it did not create
            'Me.CheckForIllegalCrossThreadCalls = False
            Control.CheckForIllegalCrossThreadCalls = False

            'Preset the Environment to DEV01
            cmbEnv.SelectedIndex = cmbEnv.Items.IndexOf("PROD01")

        Catch ex As Exception
            MsgBox("We experienced an error in the load()")
        End Try
    End Sub

    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click
        Dim y As Integer
        Try
            'Verify that at least one client was selected before continuing
            If CheckedListBox_Clients.CheckedItems.Count < 1 Then
                MsgBox("Please select 1 or more clients to continue.")
                Exit Sub
            End If

            '************************************************************

            StartedAt = Now

            objRx = CreateObject("PCOMM.autECLPS")
            objWait = CreateObject("PCOMM.autECLOIA")
            objMgr = CreateObject("PCOMM.autECLConnMgr")
            autECLConnList = CreateObject("PCOMM.autECLConnList")

            'Username
            GetUsername()
            'objWorkbook1.Worksheets(1).Cells(1, 2).Value = usrNm
            objWorkbook1.Worksheets(1).range("ReportRunBy").Value = usrNm

            'DateRange
            'objWorkbook1.Worksheets(1).Cells(2, 2).Value = Me.DateTimePickerFrom.Value & " - " & Me.DateTimePickerTo.Value
            objWorkbook1.Worksheets(1).range("DateRange").Value = Me.DateTimePickerFrom.Value & " - " & Me.DateTimePickerTo.Value


            y = ManageSessions()

            OpenNewSession()

            waitOnMe(4000)

            objMgr2 = CreateObject("PCOMM.autECLConnMgr")

            waitOnMe(1000)

            ObjSessionHandle = objMgr2.autECLConnList(y).Handle

            objRx.SetConnectionByHandle(ObjSessionHandle)
            objWait.SetConnectionByHandle(ObjSessionHandle)

            waitForMe()



            'switch the tab focus to the summary
            TabControl1.SelectedIndex = 1


            BackgroundWorker1.RunWorkerAsync()





            GetTo_JobScheduleList_Screen()

            waitForMe()

            EnterGroupNames()


            '' ''var_TimeStamp = Replace(Now, "/", "-")
            '' ''var_TimeStamp = Replace(var_TimeStamp, ":", "-")

            '' '' ''Close Results Spreadsheet
            '' ''objExcel.ActiveWorkbook.SaveAs("C:\Users\Public\Eligibility Group Rejects\EligGrpRjct_ " & var_TimeStamp & ".xlsx")


            ' '' ''****   Close Excel *********************************************************
            '' ''objExcel.ActiveWorkbook.Close()
            '' ''objExcel.Quit()
            ' '' ''****************************************************************************

            ' '' ''Clean up
            '' ''objExcel = Nothing
            '' ''objWorkbook1 = Nothing
            '' ''objWorksheet1 = Nothing
            '' ''objWorksheet2 = Nothing

            'Populate the Summary tab
            ExtractCarrierInfo()

            '*****  Close RxClaim session  **********************************************
            objMgr2.StopConnection(ObjSessionHandle)
            ''***************************************************************************

            ' BackgroundWorker1.Dispose()

            'show that we are done
            lblAllDone.Visible = True

            FinishedAt = Now

            '*******************************************************************
            Dim TotalTime As String
            Dim timeSpan As TimeSpan = FinishedAt.Subtract(StartedAt)
            Dim difHr As Integer = timeSpan.Hours
            Dim difMin As Integer = timeSpan.Minutes
            Dim difSec As Integer = timeSpan.Seconds

            TotalTime = difHr & ":" & difMin & ":" & difSec

            objWorkbook1.Worksheets(1).range("StartedAt").Value = StartedAt
            objWorkbook1.Worksheets(1).range("FinishedAt").Value = FinishedAt
            objWorkbook1.Worksheets(1).range("TotalTime").Value = TotalTime
            '*******************************************************************

            'objWorkbook1.Worksheets(1).range("ClientSummary").Value = Clients

            Dim row As Integer = objWorkbook1.Worksheets(1).range("ClientSummary").row
            Dim z As Integer = 0
            For Each item In Clients
                objWorkbook1.Worksheets(1).Cells(row + z, 1).Value = item.nm
                objWorkbook1.Worksheets(1).Cells(row + z, 2).Value = item.elpTime

                z = z + 1
            Next
        Catch ex As Exception
            MsgBox("Experienced an exception on btnStart_Click():  " & ex.ToString)
        End Try
    End Sub

    Private Sub cbAll_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles cbAll.CheckedChanged
        Try
            'if its checked...then check everything in the Checkbox List...otherwise uncheck everything
            For i = 0 To CheckedListBox_Clients.Items.Count - 1
                If cbAll.Checked Then
                    CheckedListBox_Clients.SetItemChecked(i, True)
                Else
                    CheckedListBox_Clients.SetItemChecked(i, False)
                End If
            Next
        Catch ex As Exception
            MsgBox("Experienced an exception on cbAll_CheckedChanged():  " & ex.ToString)
        End Try
    End Sub

    Private Sub Form1_FormClosed(sender As System.Object, e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            If objExcel.ActiveWorkbook Is Nothing Then
                'Do not do anything
            Else
                'Close down the spreadsheet
                objExcel.ActiveWorkbook.Close()
                objExcel = Nothing
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on closing the form:  " & ex.ToString)
        End Try
    End Sub

    Private Sub DateTimePickerFrom_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePickerFrom.ValueChanged
        If Me.DateTimePickerFrom.Value > Me.DateTimePickerTo.Value Then
            Me.DateTimePickerTo.Value = Me.DateTimePickerFrom.Value
        End If
    End Sub

    Private Sub DateTimePickerTo_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePickerTo.ValueChanged
        If Me.DateTimePickerTo.Value < Me.DateTimePickerFrom.Value Then
            Me.DateTimePickerFrom.Value = Me.DateTimePickerTo.Value
        End If
    End Sub

    '*********************************************************************************************************************************************************************
    'Functions

    Public Function ManageSessions()
        Dim intSessions, x, y As Integer

        Try
            intSessions = objMgr.autECLConnList.Count

            '** So...if we have session 1 (a), 2, (b) open and I close session 1 (a)...and now I want to open a new session...
            'the new session will be A...which is 1...that is the one I want to use.
            If intSessions > 0 Then

                For x = 1 To intSessions
                    y = 0

                    If x = 1 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "a" Then
                        y = 1
                    ElseIf x = 2 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "b" Then
                        y = 2
                    ElseIf x = 3 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "c" Then
                        y = 3
                    ElseIf x = 4 And LCase(CStr(objMgr.autECLConnList(x).Name)) <> "d" Then
                        y = 4
                    ElseIf x > 4 Then
                        'shouldn't have more than 5 sessions open...right?!
                        MsgBox("Sorry you have too many RxClaim sessions open." & Chr(13) & "Please close 1 or more and try again.")
                        ManageSessions = 0
                        Exit Function
                    End If

                    If y > 0 Then
                        Exit For
                    End If
                Next

                If y = 0 Then y = intSessions + 1

            ElseIf intSessions = 0 Then
                y = 1
            Else
                MsgBox("SOMETHING IS WRONG...")
                ManageSessions = 0
                Exit Function
            End If

            ManageSessions = y
        Catch ex As Exception
            ManageSessions = Nothing
            MsgBox("Experienced an exception on ManageSessions():  " & ex.ToString)
        End Try
    End Function

    Public Function GetPercentage(GroupRec As Integer, ByVal GroupErr As Integer) As String
        Try
            If IsNumeric(GroupRec) Then
                If GroupRec > 0 Then
                    GetPercentage = GroupErr / GroupRec
                Else
                    GetPercentage = "n/a"
                End If
            Else
                GetPercentage = ""
            End If
        Catch ex As Exception
            GetPercentage = Nothing
            MsgBox("Experienced an exception on GetPercentage():  " & ex.ToString)
        End Try
    End Function

    Function getMyDocs() As String
        Dim WshShell As Object

        Try
            WshShell = CreateObject("WScript.Shell")
            getMyDocs = WshShell.SpecialFolders("Desktop") & "\"
        Catch ex As Exception
            getMyDocs = Nothing
            MsgBox("Experienced an exception on getMyDocs():  " & ex.ToString)
        End Try
    End Function

    Function CheckForPopup(txt1, row1, col1, len1, txt2, row2, col2, len2, t_o)     't_o = is in milliseconds
        Dim bFound
        Dim theTime

        Try
            theTime = 0
            bFound = False

            waitForMe()

            'Default to False (meaning that the pop-up was not detected)
            CheckForPopup = False

            Do
                objRx.Wait(10)

                theTime = theTime + 10

                If Trim(objRx.GetText(row1, col1, len1)) = txt1 Then           'This will check if the popup was found
                    CheckForPopup = True
                    bFound = True
                ElseIf Trim(objRx.GetText(row2, col2, len2)) = txt2 Then   'This will check if the other expected page was found
                    bFound = True
                End If
            Loop While bFound = False And theTime < t_o

            'This will tell us how long its taking...
            'objWorksheet1.Cells(int_Counter_Main, 4).Value = theTime

            If theTime > t_o Then
                MsgBox("stop...we timed out")
            End If
        Catch ex As Exception
            CheckForPopup = Nothing
            MsgBox("Experienced an exception on CheckForPopup():  " & ex.ToString)
        End Try
    End Function

    Function FormatDate(myDate)
        'Lets zero fill the day and month

        Try
            If IsDate(myDate) Then
                Dim m, d, y

                'm = Right("0" & DatePart("m", myDate), 2)
                m = ("0" & DatePart("m", myDate)).ToString.Substring(("0" & DatePart("m", myDate)).Length - 2, 2)
                'd = Right("0" & DatePart("d", myDate), 2)
                d = ("0" & DatePart("d", myDate)).ToString.Substring(("0" & DatePart("d", myDate)).Length - 2, 2)
                y = DatePart("yyyy", myDate)

                FormatDate = m & "-" & d & "-" & y
            Else
                FormatDate = "          "
            End If
        Catch ex As Exception
            FormatDate = Nothing
            MsgBox("Experienced an exception on FormatDate():  " & ex.ToString)
        End Try
    End Function

    Function getScrapeSummary_GreenScreenRow(cur_Row As Integer) As Integer
        Try
            If cur_Row = 24 Then
                cur_Row = 6        'This will reset it back to row 6 and add on up to 4 more rows (for the header)

                'page down to the next page
                MoveMe("roll up", 1)

                'return the new current row
                getScrapeSummary_GreenScreenRow = cur_Row
            Else
                getScrapeSummary_GreenScreenRow = cur_Row + 1
            End If
        Catch ex As Exception
            getScrapeSummary_GreenScreenRow = Nothing
            MsgBox("Experienced an exception on getScrapeSummary_GreenScreenRow():  " & ex.ToString)
        End Try
    End Function

    '*********************************************************************************************************************************************************************
    'Subs

    Public Sub OpenNewSession()
        Dim Envir As String

        Try
            'Now find the "File name" to open up based on their selection
            If LCase(cmbEnv.SelectedItem) = "dev01" Then
                Envir = "Dev01.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "dev02" Then
                Envir = "Dev02.AS4"
            ElseIf LCase(cmbEnv.SelectedItem) = "prod03" Then
                'If usrNm = "rlberg" Then
                Envir = "PROD03.AS4"
                'Else
                '   MsgBox("sorry...you cannot use PROD03 at this time...it is still in testing.")
                '  Exit Sub
                'End If
            ElseIf LCase(cmbEnv.SelectedItem) = "prod01" Then
                Envir = "PROD01.AS4"
                'MsgBox("Sorry...not available to use at this time")
                'Exit Sub
            Else
                MsgBox("Environment was not found ... exiting.")
                Exit Sub
            End If


            '***********************************************

            Dim sDir As String = getMyDocs()

            If sDir.Length > 1 Then
                'Now we are trying to open up a session
                Try
                    Process.Start(sDir & "RxClaims Sessions\" & Envir)
                Catch
                    Try
                        Process.Start("C:\Users\Public\Desktop\RxClaims Sessions\" & Envir)
                    Catch ex As Exception
                        MsgBox("Please open up an RxClaim session and then press 'OK' to this message")
                    End Try
                End Try
            Else
                MsgBox("couldn't find Desktop")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on OpenNewSession():  " & ex.ToString)
        End Try
    End Sub

    Public Sub OpenSpreadsheetTemplate()
        Try
            'objExcelFilePath = "C:\Users\rlberg\Desktop\Eligibility Group Rejects Output File.xlsx"

            'objExcelFilePath = "C:\Users\rlberg\Desktop\Stage Two - Eligibility Rejects Output File_Fixing.xlsx"

            'objExcelFilePath = "J:\shrproj\Benefit Operations\Paper Claims\Claims Processing\Eligibility Rejects Macro\Eligibility Group Rejects Output File.xlsx"
            'objExcelFilePath = "J:\shrproj\Benefit Operations\Paper Claims\Claims Processing\Eligibility Rejects Macro\Stage Two - Eligibility Group Rejects Output File.xlsx"
            'objExcelFilePath = "J:\shrproj\Benefit Operations\Paper Claims\Claims Processing\Eligibility Rejects Macro\Stage Two - Eligibility Group Rejects Output File_Ryan.xlsx"
            'objExcelFilePath = "J:\shrproj\Benefit Operations\Paper Claims\Claims Processing\Eligibility Rejects Macro\Stage Two - Eligibility Group Rejects Output File_Ryan_Medicare.xlsx"
            objExcelFilePath = "J:\shrproj\Benefit Operations\Paper Claims\Claims Processing\Eligibility Rejects Macro\Stage Two - Eligibility Rejects Output File.xlsx"

            If objExcelFilePath = Nothing Then
                MsgBox("Sorry we couldn't find that spreadsheet")
                Exit Sub
            End If

            objExcel = CreateObject("Excel.Application")
            objWorkbook1 = objExcel.Workbooks.Open(objExcelFilePath)

            objExcel.Visible = True     '--only do this if you want to see the progress
        Catch ex As Exception
            MsgBox("Experienced an exception on OpenSpreadsheetTemplate():  " & ex.ToString)
        End Try
    End Sub

    Public Sub AddCarrierInfo(ByVal Type As Integer, ByVal GroupRec As Integer, ByVal GroupRej As Integer, ByVal MemberRec As Integer, ByVal MemberRej As Integer)
        Dim i As Integer

        Try
            For i = 0 To Carriers.Count - 1
                If Carriers(i).[jn].ToString = JobName And Carriers(i).[tp].ToString = Type Then
                    Carriers(i).grpRec = Carriers(i).grpRec + GroupRec
                    Carriers(i).grpRej = Carriers(i).grpRej + GroupRej
                    Carriers(i).mbrRec = Carriers(i).mbrRec + MemberRec
                    Carriers(i).mbrRej = Carriers(i).mbrRej + MemberRej
                    Exit Sub
                End If
            Next

            'if we got this far...that means we did not find it and thus must add a new one
            Carriers.Add(New Carrier(JobName, Type, GroupRec, GroupRej, MemberRec, MemberRej))
        Catch ex As Exception
            MsgBox("Experienced an exception on AddCarrierInfo():  " & ex.ToString)
        End Try
    End Sub

    Public Sub ExtractCarrierInfo()
        Try
            If Carriers.Count > 0 Then
                Dim i, j As Integer
                Dim bFound As Boolean
                'Loop thru the Summary tab
                'Loop thru the Carriers object to

                j = 2   'start at row 2
                Do While Len(objWorkbook1.Worksheets(1).Cells(j, 3).Value) > 0
                    bFound = False  'Default to not found

                    For i = 0 To Carriers.Count - 1
                        'Column 3 is "Stage Job Name"
                        'Column 4 is "Load Job Name"
                        'Column 5 is "Recycle"  [which means tells us how many files to expect {should be 1 or 2} ]

                        If Carriers(i).tp = 1 Then        'Does NOT Recycle and should only have 1 file per day
                            'Column 3 is "Stage Job Name"  ...We care about both Total Records and Rejects
                            If Trim(objWorkbook1.Worksheets(1).Cells(j, 3).Value) = Carriers(i).[jn].ToString Then
                                'Member
                                objWorkbook1.Worksheets(1).Cells(j, 7).Value = objWorkbook1.Worksheets(1).Cells(j, 7).Value + Carriers(i).mbrRec   'Total Records  (For JUST Load)
                                objWorkbook1.Worksheets(1).Cells(j, 8).Value = objWorkbook1.Worksheets(1).Cells(j, 8).Value + Carriers(i).mbrRej   'Total Rejects   (for Stage and Load)
                                objWorkbook1.Worksheets(1).Cells(j, 9).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 7).Value, objWorkbook1.Worksheets(1).Cells(j, 8).Value)

                                'Group
                                objWorkbook1.Worksheets(1).Cells(j, 11).Value = objWorkbook1.Worksheets(1).Cells(j, 11).Value + Carriers(i).grpRec       'Total Records  (For JUST Load)
                                objWorkbook1.Worksheets(1).Cells(j, 12).Value = objWorkbook1.Worksheets(1).Cells(j, 12).Value + Carriers(i).grpRej       'Total Rejects   (for Stage and Load)
                                objWorkbook1.Worksheets(1).Cells(j, 13).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 11).Value, objWorkbook1.Worksheets(1).Cells(j, 12).Value)

                                bFound = True
                            End If

                            'Column 4 is "Load Job Name" ...All we care about is Rejects
                            If Trim(objWorkbook1.Worksheets(1).Cells(j, 4).Value) = Carriers(i).[jn].ToString Then
                                'Member
                                objWorkbook1.Worksheets(1).Cells(j, 8).Value = objWorkbook1.Worksheets(1).Cells(j, 8).Value + Carriers(i).mbrRej   'Total Rejects   (for Stage and Load)
                                objWorkbook1.Worksheets(1).Cells(j, 9).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 7).Value, objWorkbook1.Worksheets(1).Cells(j, 8).Value)

                                'Group
                                objWorkbook1.Worksheets(1).Cells(j, 12).Value = objWorkbook1.Worksheets(1).Cells(j, 12).Value + Carriers(i).grpRej       'Total Rejects   (for Stage and Load)
                                objWorkbook1.Worksheets(1).Cells(j, 13).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 11).Value, objWorkbook1.Worksheets(1).Cells(j, 12).Value)

                                bFound = True
                            End If
                        ElseIf Carriers(i).tp = 2 And (i + 3) <= Carriers.Count Then    'This would be the "RECYCLED" file
                            'if this record has the type set to 2 then the record direct behind it should be set to 1 (original)

                            Dim iStageOrigInputGroup, iStageOrigInputMember As Integer              'A
                            Dim iStageOrigRejGroup, iStageOrigRejMember As Integer                  'B
                            Dim iStageRecycleInputGroup, iStageRecycleInputMember As Integer        'C
                            Dim iLoadOrigRejectGroup, iLoadOrigRejectMember As Integer              'D
                            Dim iLoadRecyleInputGroup, iLoadRecyleInputMember As Integer            'E
                            Dim iStageRejectRecycleGroup, iStageRejectRecycleMember As Integer      'F
                            Dim iLoadRejectRecycleGroup, iLoadRejectRecycleMember As Integer        'G

                            'Total Records = A
                            '
                            'Total Rejects = (B-C + F) + (D-E + G)

                            '*      (Stage REJ ORG - Stage Input Rec + Stage REJ Rec) + (Load Rej ORG - Load Input Recyle + Load REJ Recycle)
                            '*
                            '*      If (Stage REJ ORG - Stage Input Rec) < 0 … then make it equal to zero
                            '*      If (Load Rej ORG - Load Input Recyle) < 0 … then make it equal to zero


                            bFound = False  'Default to not found (ONLY DOING THIS FOR THE "RECYCLED")...otherwise this is handled just after the "Do-While" loop above

                            'If i > 0 Then
                            '    j = j + 1       'Move the row counter for Excel to the next row
                            'End If

                            'Stage
                            If Trim(objWorkbook1.Worksheets(1).Cells(j, 3).Value) = Carriers(i).[jn].ToString Then
                                If Carriers(i + 1).tp = 1 And (Carriers(i).jn = Carriers(i + 1).jn) Then    'if the next record is the same job name...but it is the Orginial...then we can move forward
                                    'A
                                    iStageOrigInputGroup = Carriers(i + 1).grpRec
                                    iStageOrigInputMember = Carriers(i + 1).mbrRec
                                    'B
                                    iStageOrigRejGroup = Carriers(i + 1).grpRej
                                    iStageOrigRejMember = Carriers(i + 1).mbrRej
                                    'C
                                    iStageRecycleInputGroup = Carriers(i).grpRec
                                    iStageRecycleInputMember = Carriers(i).mbrRec
                                    'F
                                    iStageRejectRecycleGroup = Carriers(i).grpRej
                                    iStageRejectRecycleMember = Carriers(i).mbrRej

                                    bFound = True
                                End If
                            End If

                            'Load
                            If Trim(objWorkbook1.Worksheets(1).Cells(j, 4).Value) = Carriers(i + 2).[jn].ToString Then
                                If Carriers(i + 3).tp = 1 And (Carriers(i + 2).jn = Carriers(i + 3).jn) Then    'if the next record is the same job name...but it is the Orginial...then we can move forward
                                    'D
                                    iLoadOrigRejectGroup = Carriers(i + 3).grpRej
                                    iLoadOrigRejectMember = Carriers(i + 3).mbrRej
                                    'E
                                    iLoadRecyleInputGroup = Carriers(i + 2).grpRec
                                    iLoadRecyleInputMember = Carriers(i + 2).mbrRec
                                    'G
                                    iLoadRejectRecycleGroup = Carriers(i + 2).grpRej
                                    iLoadRejectRecycleMember = Carriers(i + 2).mbrRej

                                    bFound = True
                                End If
                            End If

                            If bFound = True Then
                                Dim Group1, Group2, Member1, Member2 As Integer

                                Group1 = iStageOrigRejGroup - iStageRecycleInputGroup
                                If Group1 < 0 Then Group1 = 0

                                Group2 = iLoadOrigRejectGroup - iLoadRecyleInputGroup
                                If Group2 < 0 Then Group2 = 0

                                Member1 = iStageOrigRejMember - iStageRecycleInputMember
                                If Member1 < 0 Then Member1 = 0

                                Member2 = iLoadOrigRejectMember - iLoadRecyleInputMember
                                If Member2 < 0 Then Member2 = 0

                                'Group
                                'objWorkbook1.Worksheets(1).Cells(j, 7).Value = iStageOrigInputGroup
                                'objWorkbook1.Worksheets(1).Cells(j, 8).Value = (Group1 + iStageRejectRecycleGroup) + (Group2 + iLoadRejectRecycleGroup)
                                'objWorkbook1.Worksheets(1).Cells(j, 9).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 7).Value, objWorkbook1.Worksheets(1).Cells(j, 8).Value)
                                objWorkbook1.Worksheets(1).Cells(j, 11).Value = iStageOrigInputGroup
                                objWorkbook1.Worksheets(1).Cells(j, 12).Value = (Group1 + iStageRejectRecycleGroup) + (Group2 + iLoadRejectRecycleGroup)
                                objWorkbook1.Worksheets(1).Cells(j, 13).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 11).Value, objWorkbook1.Worksheets(1).Cells(j, 12).Value)

                                'Member
                                'objWorkbook1.Worksheets(1).Cells(j, 11).Value = iStageOrigInputMember
                                'objWorkbook1.Worksheets(1).Cells(j, 12).Value = (Member1 + iStageRejectRecycleMember) + (Member2 + iLoadRejectRecycleMember)
                                'objWorkbook1.Worksheets(1).Cells(j, 13).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 11).Value, objWorkbook1.Worksheets(1).Cells(j, 12).Value)
                                objWorkbook1.Worksheets(1).Cells(j, 7).Value = iStageOrigInputMember
                                objWorkbook1.Worksheets(1).Cells(j, 8).Value = (Member1 + iStageRejectRecycleMember) + (Member2 + iLoadRejectRecycleMember)
                                objWorkbook1.Worksheets(1).Cells(j, 9).Value = GetPercentage(objWorkbook1.Worksheets(1).Cells(j, 7).Value, objWorkbook1.Worksheets(1).Cells(j, 8).Value)

                                i = i + 3       'If we made it here ... our Carriers object will work with 4 at a time


                                'Else
                                '    'we need to add "n/a" to that carrier
                                '    objWorkbook1.Worksheets(1).Cells(j, 9).Value = "n/a"
                                '    objWorkbook1.Worksheets(1).Cells(j, 13).Value = "n/a"
                            End If
                        End If
                    Next

                    If bFound = False And Len(Trim(objWorkbook1.Worksheets(1).Cells(j, 9).Value)) < 1 Then
                        'we need to add "n/a" to that carrier
                        objWorkbook1.Worksheets(1).Cells(j, 9).Value = "n/a"
                        objWorkbook1.Worksheets(1).Cells(j, 13).Value = "n/a"
                    End If



                    'Reformat columns 9 and 13 to "Percent"



                    j = j + 1
                Loop
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on ExtractCarrierInfo():  " & ex.ToString)
        End Try
    End Sub

    Public Sub GetTo_JobScheduleList_Screen()
        'IF 19,2 for 11 = "Press Enter"  ...  This is usually the 1st screen that shows if you already have another session open
        If Trim(objRx.GetText(19, 2, 11)) = "Press Enter" Then
            objRx.SendKeys("[Enter]")
            waitOnMe(1000)
        End If

        'Now try connecting to that session ... we will wait 7 seconds
        'This is a hard wait to ensure that the RxClaim session has started
        IsRightScreenName("QPADEV", 1, 70, 5000)

        waitForMe()

        If LCase(cmbEnv.SelectedItem) = "prod03" Then
            'objRx.SetText("PPF", 21, 7)
            MsgBox("RBO-Robot Main Menu library is not available in/for PROD03")
            Exit Sub
        Else
            objRx.SetText("RBO", 21, 7)
        End If

        waitForMe()
        MoveMe("enter", 1)
        waitForMe()

        IsRightScreenName("RSL201", 1, 2, 5000)
        waitForMe()
        objRx.SetText("1", 21, 18)
        waitForMe()
        MoveMe("enter", 1)
        waitForMe()

        IsRightScreenName("RBT1010", 1, 2, 5000)
        waitForMe()
        objRx.SetText("1", 21, 36)
        waitForMe()
        MoveMe("enter", 1)
        waitForMe()
    End Sub

    Public Sub PopulateClientList()
        'Walk thru each tab name starting with the 2nd tab and populate the checkbox list
        'For i = 2 To objWorkbook1.Worksheets.Count
        '    CheckedListBox_Clients.Items.Add(objWorkbook1.Worksheets(i).name)
        'Next

        'Create list and add all values
        Dim values As List(Of String) = New List(Of String)
        Dim s As String

        For i = 2 To objWorkbook1.Worksheets.Count
            s = objWorkbook1.Worksheets(i).name
            values.Add(Trim(s.Substring(0, s.IndexOf("-"))))        'We are expecting the format to be something like -->  "MT - G" or "MT - M"...we just want to add "MT"
        Next

        'Filter distinct values and add into a new list
        Dim result As List(Of String) = values.Distinct().ToList

        'Add new distinct value to checkedListBox
        For Each element As String In result
            CheckedListBox_Clients.Items.Add(element)
        Next
    End Sub

    Public Sub EnterGroupNames()
        Try
            '1st tab on spreadsheet should be Summary
            'As of right now we have 17 clients

            'DO THIS TO GET THE SEARCH RESULTS CORRECT
            'F9...Job Search Criteria
            waitForMe()
            MoveMe("pf9", 1)
            waitForMe()

            'Enter "1" for "Job Name"
            '*This is because this Popup can/does open up in different locations!!!  UGH!  ********************************************
            If LCase(Trim(objRx.GetText(12, 38, 8))) = "job name" Then
                objRx.SetText("1", 12, 34)
            ElseIf LCase(Trim(objRx.GetText(12, 40, 8))) = "job name" Then
                objRx.SetText("1", 12, 36)
            ElseIf LCase(Trim(objRx.GetText(12, 41, 8))) = "job name" Then
                objRx.SetText("1", 12, 37)
            ElseIf LCase(Trim(objRx.GetText(12, 43, 8))) = "job name" Then
                objRx.SetText("1", 12, 39)
            ElseIf LCase(Trim(objRx.GetText(12, 46, 8))) = "job name" Then
                objRx.SetText("1", 12, 42)
            ElseIf LCase(Trim(objRx.GetText(12, 47, 8))) = "job name" Then
                objRx.SetText("1", 12, 43)
            ElseIf LCase(Trim(objRx.GetText(16, 8, 8))) = "job name" Then
                objRx.SetText("1", 16, 4)
            End If
            '**************************************************************************************************************************

            waitForMe()

            ''Hit Enter
            MoveMe("enter", 1)
            waitForMe()

            'Default
            int_TabNum = 0

            Dim iCounter As Integer = 0

            For Each item In Me.CheckedListBox_Clients.CheckedItems

                '******************************************************************************
                Dim TimeStartedClient As DateTime = Now
                Dim TimeEndedClient As DateTime
                '******************************************************************************

                'Go thru once for Group and once for Member
                Me.lblClientName.Text = item.ToString()

                iCounter = iCounter + 1

                Me.lblClientStatus.Text = iCounter & " of " & Me.CheckedListBox_Clients.CheckedItems.Count

                ClearResultsLabels()

                'Group
                'We are adding 2 because the checkbox list starts at zero and the spreadsheet has the 1st tab as Summary...so starting with tab #2
                int_TabNum = (CheckedListBox_Clients.Items.IndexOf(item) + 1) * 2
                SSTabType = "G"
                Me.lblGroupStatus.Text = "Running..."
                EnterJobName()
                Me.lblGroupStatus.Text = "Finished"

                'Update Progress Bar  *************************************************************************************
                Me.ProgressBar_Results.Value = ((iCounter - 0.5) / Me.CheckedListBox_Clients.CheckedItems.Count) * 100
                '**********************************************************************************************************

                'Member
                int_TabNum = int_TabNum + 1     'the following tab will be the Member
                SSTabType = "M"
                Me.lblMemberStatus.Text = "Running..."
                EnterJobName()
                Me.lblMemberStatus.Text = "Finished"

                'Update Progress Bar  *************************************************************************************
                Me.ProgressBar_Results.Value = (iCounter / Me.CheckedListBox_Clients.CheckedItems.Count) * 100
                '**********************************************************************************************************

                '*******************************************************************
                TimeEndedClient = Now

                Dim TotalTime As String
                Dim timeSpan As TimeSpan = TimeEndedClient.Subtract(TimeStartedClient)
                Dim difHr As Integer = timeSpan.Hours
                Dim difMin As Integer = timeSpan.Minutes
                Dim difSec As Integer = timeSpan.Seconds

                TotalTime = difHr & ":" & difMin & ":" & difSec

                Clients.Add(New Client(item.ToString(), TotalTime))
                '*******************************************************************
            Next
        Catch ex As Exception
            MsgBox("Experienced an exception on EnterGroupNames():  " & ex.ToString)
        End Try
    End Sub

    Sub ClearResultsLabels()
        Me.lblGroupStatus.Text = ""
        Me.lblMemberStatus.Text = ""
    End Sub

    Sub EnterJobName()
        Try
            'Reset this global varaible back to 4...this will mean that each time we are in a new tab...default the row to 4
            'SSRowNumber = 4

            SSRowNumber = 1

            'This will let us traverse thru each of the jobs on a tab
            Do While SSRowNumber < objWorkbook1.Worksheets(int_TabNum).UsedRange.rows.count
                'search for the JobName
                waitForMe()
                IsRightScreenName("RBT276", 1, 2, 1000)   'was 5000

                waitForMe()

                'This will clear out the field so we can enter new/clean data
                MoveMe2("eraseeof", 5, 32)

                objRx.SetText(Trim(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 3).Value), 5, 32)
                waitForMe()
                MoveMe("enter", 1)
                waitForMe()

                '****
                JobName = Trim(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 3).Value)
                JobScheduleList()

                'Go back to the RBT76 Screen
                Do Until (objRx.WaitForString("RBT276", 1, 2, 1000, True))   'was 5000
                    waitForMe()
                    MoveMe("pf3", 1)
                    waitForMe()
                Loop

                'Move to the next job on the Excel tab
                SSRowNumber = SSRowNumber + 1
            Loop
        Catch ex As Exception
            MsgBox("Experienced an exception on EnterJobName():  " & ex.ToString)
        End Try
    End Sub

    Sub JobScheduleList()
        Dim bFoundJob As Boolean = False
        'Find and select the Job Name that we previously search for

        Try
            'starts with 9,18
            'we are stopping with 20,18

            'this will put us on the row where we will do an INSERT
            If objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber + 3, 1).Value = "" Then
                SSRowNumber = SSRowNumber + 2       'This would be if we are doing the last one on the tab...(we don't want to have too many rows inserted)
            Else
                SSRowNumber = SSRowNumber + 3
            End If

            For i = 9 To 20
                If Trim(objRx.GetText(i, 18, 10)) = Trim(JobName) Then
                    'MsgBox("we found it")
                    bFoundJob = True
                    objRx.SetText("11", i, 2)
                    waitForMe()
                    MoveMe("enter", 1)
                    waitForMe()

                    JobCompletionHistory()

                    Exit Sub
                End If
            Next

            'If the job was not found...make note of it
            If bFoundJob = False Then
                objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()

                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Job was not found:"
                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & Trim(JobName)

                SSRowNumber = SSRowNumber + 1
                objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on JobScheduleList():  " & ex.ToString)
        End Try
    End Sub


    Public Function FindJobByDate(dateToSearch As String) As Integer
        Dim scrapeValue As String
        Dim atLeastOneWasFound As Boolean
        Dim iTimesPagedDown As Integer = 0
        Dim iPreFoundCount As Integer = 0
        Dim iPostFoundCount As Integer = 0
        Dim iHowManyPagesWeCanUse As Integer = 5        'Constant

        '************************************************************************************************************
        'default
        atLeastOneWasFound = False
        iPreFoundCount = 0
        iPostFoundCount = 0
        '************************************************************************************************************

        For i = 8 To 19     'We are presuming that everything we will need will be on the 1st page
            scrapeValue = Trim(objRx.GetText(i, 50, 5))

            If scrapeValue.Length = 0 Then Exit For

            If dateToSearch = scrapeValue Then
                iPreFoundCount = iPreFoundCount + 1
                atLeastOneWasFound = True
            End If

            If i = 19 And atLeastOneWasFound = False And iTimesPagedDown < iHowManyPagesWeCanUse Then       'Page down and start over
                MoveMe("roll up", 1)    'Page down
                iTimesPagedDown = iTimesPagedDown + 1
                i = 7   'Reset the for loop to (8-1)
            End If
        Next

        'Make sure we are 'paged' all the way up...this could happen if we were previously here and paged down...it would then return to the same spot...we want to go all the we back to the 1st (top) result
        Dim z As Integer

        For z = 1 To iTimesPagedDown
            MoveMe("roll down", 1)    'Page down
        Next

        FindJobByDate = iPreFoundCount
    End Function


    Sub JobCompletionHistory()
        'Find the job(s) that are within the date range specified
        'Enter "4"  and ENTER
        Dim dateToSearch As String
        Dim toDate As Date      'Actual Date from UI
        Dim fromDate As Date    'Actual Date from UI
        'Dim atLeastOneWasFound As Boolean
        Dim scrapeValue As String
        Dim iCounter As Integer = 0

        'Dim iRecyleNum As Integer = 0
        Dim iRowToSelect As Integer = 0
        Dim iPreFoundCount As Integer = 0
        Dim iPostFoundCount As Integer = 0
        Dim iHowManyTimesThru As Integer = 0
        Dim iTimesPagedDown As Integer = 0

        Dim iHowManyPagesWeCanUse As Integer = 5        'Constant
        Dim iDateRangeCount As Integer                  'Tells us how many days are in our date range

        Try
            fromDate = Me.DateTimePickerFrom.Value.Date
            toDate = Me.DateTimePickerTo.Value.Date

            iDateRangeCount = DateDiff(DateInterval.Day, fromDate, toDate)      'if it is zero...it means that they are the same day

            dateToSearch = fromDate.Month.ToString & "/" & Microsoft.VisualBasic.Right("0" & fromDate.Day.ToString, 2)

            Do Until fromDate > toDate


                iPreFoundCount = FindJobByDate(dateToSearch)

                If iPreFoundCount = 0 And iDateRangeCount = 0 Then
                    'Then grab results from the previous day
                    fromDate = fromDate.AddDays(-1)
                    dateToSearch = fromDate.Month.ToString & "/" & Microsoft.VisualBasic.Right("0" & fromDate.Day.ToString, 2)
                    fromDate = fromDate.AddDays(1) 'set it back to what it was
                    iPreFoundCount = FindJobByDate(dateToSearch)

                    If iPreFoundCount > 0 Then
                        'FIND THE RIGHT ROW ON THE SUMMARY TAB
                        Dim k, iRecNum As Integer

                        For k = 2 To objWorkbook1.Worksheets(1).UsedRange.rows.count
                            If JobName = Trim(objWorkbook1.Worksheets(1).Cells(k, 3).Value) Or JobName = Trim(objWorkbook1.Worksheets(1).Cells(k, 4).Value) Then
                                iRecNum = k     'all we want is the row # that this job is on
                                Exit For
                            End If
                        Next

                        If iRecNum > 0 Then
                            objWorkbook1.Worksheets(1).Cells(iRecNum, 2).Value = "*" & dateToSearch
                        End If
                    End If
                End If





                ''Make sure we are 'paged' all the way up...this could happen if we were previously here and paged down...it would then return to the same spot...we want to go all the we back to the 1st (top) result
                'Dim z As Integer

                'For z = 1 To iTimesPagedDown
                '    MoveMe("roll down", 1)    'Page down
                'Next
                ''************************************************************************************************************
                ''default
                'atLeastOneWasFound = False
                'iPreFoundCount = 0
                'iPostFoundCount = 0
                ''************************************************************************************************************

                'For i = 8 To 19     'We are presuming that everything we will need will be on the 1st page
                '    scrapeValue = Trim(objRx.GetText(i, 50, 5))

                '    If scrapeValue.Length = 0 Then Exit For

                '    If dateToSearch = scrapeValue Then
                '        iPreFoundCount = iPreFoundCount + 1
                '        atLeastOneWasFound = True
                '    End If

                '    If i = 19 And atLeastOneWasFound = False And iTimesPagedDown < iHowManyPagesWeCanUse Then       'Page down and start over
                '        MoveMe("roll up", 1)    'Page down
                '        iTimesPagedDown = iTimesPagedDown + 1
                '        i = 7   'Reset the for loop to (8-1)
                '    End If
                'Next







                If Trim(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value).Length < 1 Then
                    SSRowNumber = SSRowNumber + 1
                End If

                If iPreFoundCount < 1 Then
                    'we need to indicate that we did NOT find a match
                    objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
                    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  No job loaded for this date...we did not find any date matches"

                    SSRowNumber = SSRowNumber + 1
                    objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
                    'ElseIf iPreFoundCount > 2 Then
                    '    'we need to indicate that we did NOT find a match
                    '    objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
                    '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                    '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  No job loaded for this date...we found more than 2 date matches"

                    '    SSRowNumber = SSRowNumber + 1
                    '    objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
                Else

                    For i = 8 To 19     'We are presuming that everything we will need will be on the 1st page
                        If iCounter > 0 Then
                            'Go back to the RBT78 Screen
                            Do Until (objRx.WaitForString("RBT278", 1, 2, 1000, True))   'was 5000
                                waitForMe()
                                MoveMe("pf3", 1)
                                waitForMe()
                            Loop
                        End If

                        scrapeValue = Trim(objRx.GetText(i, 50, 5))

                        If dateToSearch = scrapeValue Then
                            Dim k, iRecNum As Integer

                            'New...added 12/17/15   ...   (This will put separation between multiple results per Job Name) (if i=8 that means it is the 1st one in the group)

                            'If SSRowNumber > 4 Then
                            'If LCase(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 2, 1).Value) = "carrier" Or LCase(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber + 2, 1).Value) = "carrier" Then
                            'If LCase(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber + 2, 1).Value) = "carrier" Then
                            '    SSRowNumber = SSRowNumber + 1
                            'End If
                            If Len(Trim(objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 1).Value)) > 0 Then
                                SSRowNumber = SSRowNumber + 1
                            End If
                            'End If

                            objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()

                            'If iPreFoundCount = 1 Then
                            '    iRowToSelect = i
                            '    Me.iCarrierType = 1 'Original
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  Original"
                            'ElseIf iPreFoundCount = 2 And iPostFoundCount = 1 Then
                            '    iRowToSelect = i
                            '    Me.iCarrierType = 2 'Recycle
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  Recycled"
                            'ElseIf iPreFoundCount = 2 And iPostFoundCount = 2 Then

                            '    SSRowNumber = SSRowNumber + 1

                            '    iRowToSelect = i
                            '    Me.iCarrierType = 1 'Original
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  Original"
                            'Else
                            '    objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Extra File..."
                            '    Exit For
                            'End If

                            For k = 2 To objWorkbook1.Worksheets(1).UsedRange.rows.count
                                If JobName = Trim(objWorkbook1.Worksheets(1).Cells(k, 3).Value) Or JobName = Trim(objWorkbook1.Worksheets(1).Cells(k, 4).Value) Then
                                    iRecNum = objWorkbook1.Worksheets(1).Cells(k, 5).Value
                                    Exit For
                                End If
                            Next

                            'If we already have been thru here twice...reset it (this is needed when doing this for multiple days)
                            If (iRecNum = 2 And iPostFoundCount = 2) Or (iRecNum = 1 And iPostFoundCount = 1) Then
                                iPostFoundCount = 0
                            End If

                            iPostFoundCount = iPostFoundCount + 1

                            '**********************************************************

                            Dim bMoreFiles As Boolean = False

                            If iRecNum = 1 And iPostFoundCount = 1 Then
                                iRowToSelect = i
                                Me.iCarrierType = 1 'Original
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  Original"
                            ElseIf iRecNum = 2 And iPostFoundCount = 1 Then
                                iRowToSelect = i
                                Me.iCarrierType = 2 'Recycle
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  Recycled"
                            ElseIf iRecNum = 2 And iPostFoundCount = 2 Then

                                SSRowNumber = SSRowNumber + 1

                                iRowToSelect = i
                                Me.iCarrierType = 1 'Original
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Date:"
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = "'" & dateToSearch & "  --  Original"
                            Else  'it is more files than we expected...
                                iRowToSelect = i        'new as of 2/16/15
                                SSRowNumber = SSRowNumber + 1
                                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "Extra File..."
                                SSRowNumber = SSRowNumber + 1
                                objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
                                bMoreFiles = True
                            End If

                            If bMoreFiles = False Then
                                'add 4
                                objRx.SetText("4", i, 2)
                                MoveMe("enter", 1)

                                SSRowNumber = SSRowNumber + 1
                                objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()

                                JobSpooledFiles()
                            End If

                            iCounter = iCounter + 1
                        End If
                    Next
                End If

                '***************************************************************
                iHowManyTimesThru = iHowManyTimesThru + 1
                fromDate = fromDate.AddDays(1)
                dateToSearch = fromDate.Month.ToString & "/" & Microsoft.VisualBasic.Right("0" & fromDate.Day.ToString, 2)

            Loop
        Catch ex As Exception
            MsgBox("Experienced an exception on JobCompletionHistory():  " & ex.ToString)
        End Try
    End Sub

    Sub JobSpooledFiles()
        Dim scrapeValue As String
        Dim bFileFound As Boolean = False
        Dim sFileToFind As String = ""

        Try
            'Do the Group 1st
            For i = 11 To 20     'We are presuming that everything we will need will be on the 1st page
                scrapeValue = Trim(objRx.GetText(i, 7, 10))

                'IF GROUP
                If SSTabType = "G" Then
                    sFileToFind = "RXELGGER"
                    If scrapeValue.ToUpper = sFileToFind Then
                        'add 4
                        waitForMe()
                        objRx.SetText("5", i, 3)
                        waitForMe()
                        MoveMe("enter", 1)

                        ScrapeSpooledFile_Group()
                        bFileFound = True
                        Exit For        'there should only be one "RXELGGER"
                    End If
                ElseIf SSTabType = "M" Then
                    sFileToFind = "RXELGMER"
                    If scrapeValue.ToUpper = sFileToFind Then
                        'add 4
                        waitForMe()
                        objRx.SetText("5", i, 3)
                        waitForMe()
                        MoveMe("enter", 1)

                        ScrapeSpooledFile_Member()
                        bFileFound = True
                        Exit For        'there should only be one "RXELGGER"
                    End If
                End If
            Next

            If bFileFound = False Then
                objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = "  **  Could not find this file:  " & sFileToFind

                SSRowNumber = SSRowNumber + 1
                objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on JobSpooledFiles():  " & ex.ToString)
        End Try
    End Sub

    Sub ScrapeSpooledFile_Group()
        'Lets start scraping!!!! 
        'Dim ssRow As Integer = 4        'SpreadSheet Row
        Dim gsRow As Integer = 6        'Green Screen Row
        Dim bDoneScraping As Boolean = False

        Try
            Do Until bDoneScraping = True

                If UCase(Trim(objRx.GetText(gsRow, 68, 14))) = "R X  C L A I M" Then
                    gsRow = gsRow + 4

                    If gsRow > 24 Then

                        gsRow = (gsRow - 25) + 6        'This will reset it back to row 6 and add on up to 4 more rows (for the header)

                        'page down to the next page
                        MoveMe("roll up", 1)

                        'loop to the next one
                        Continue Do
                    End If
                End If

                If LCase(Trim(objRx.GetText(gsRow, 3, 8))) = "carrier:" Then
                    'get the summary info
                    ScrapeSummary_Group(int_TabNum, gsRow)

                    bDoneScraping = True
                Else
                    Dim errField1, errField2, errField3, errField4 As String
                    Dim bAppend As Boolean = False

                    'If the length of this is zero then we are appending to what we already have....
                    If Len(Trim(objRx.GetText(gsRow, 3, 10))) < 1 Then
                        bAppend = True
                    End If

                    If bAppend = False Then
                        'Insert as we go...
                        objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()

                        'Carrier
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsRow, 3, 10))

                        'Account
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsRow, 13, 15))

                        'Group
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 3).Value = Trim(objRx.GetText(gsRow, 29, 15))

                        'Group Name
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 4).Value = Trim(objRx.GetText(gsRow, 45, 40)) & " " & Trim(objRx.GetText(gsRow + 1, 45, 40))

                        'Error Messages
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 5).Value = Trim(objRx.GetText(gsRow, 86, 35)) & " " & Trim(objRx.GetText(gsRow + 1, 86, 35))

                        'Error Field
                        errField1 = Trim(objRx.GetText(gsRow, 121, 12))

                        If gsRow + 1 = 25 And Trim(objRx.GetText(gsRow + 1, 121, 12)) = "More..." Then
                            errField3 = ""
                        Else
                            errField3 = Trim(objRx.GetText(gsRow + 1, 121, 12))
                        End If

                        'F20
                        waitForMe()
                        MoveMe("pf20", 1)
                        waitForMe()

                        errField2 = Trim(objRx.GetText(gsRow, 2, 20))
                        errField4 = Trim(objRx.GetText(gsRow + 1, 2, 20))

                        'f19
                        waitForMe()
                        MoveMe("pf19", 1)
                        waitForMe()

                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 6).Value = errField1 & " " & errField2 & " " & errField3 & " " & errField4

                        SSRowNumber = SSRowNumber + 1

                        If gsRow = 23 Or gsRow = 24 Then
                            gsRow = 6   'Reset to the beginning

                            'page down to the next page
                            MoveMe("roll up", 1)
                        Else
                            If Len(Trim(objRx.GetText(gsRow, 3, 8))) > 0 Then
                                gsRow = gsRow + 2
                            Else
                                gsRow = gsRow + 1
                            End If
                        End If

                    Else
                        'This means it is carry-over from the page before...

                        'Group Name
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 4).Value = objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 4).Value & " " & Trim(objRx.GetText(gsRow, 45, 40))

                        'Error Messages
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 5).Value = objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 5).Value & " " & Trim(objRx.GetText(gsRow, 86, 35))

                        'Error Field
                        errField1 = Trim(objRx.GetText(gsRow, 121, 12))

                        If gsRow + 1 = 25 And Trim(objRx.GetText(gsRow + 1, 121, 12)) = "More..." Then
                            errField3 = Trim(objRx.GetText(gsRow + 1, 121, 12))
                        Else
                            errField3 = ""
                        End If

                        'F20
                        waitForMe()
                        MoveMe("pf20", 1)
                        waitForMe()

                        errField2 = Trim(objRx.GetText(gsRow, 2, 20))

                        'f19
                        waitForMe()
                        MoveMe("pf19", 1)
                        waitForMe()

                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 6).Value = objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 6).Value & " " & errField1 & " " & errField2

                        'If gsRow = 23 Or gsRow = 24 Then
                        If gsRow = 24 Then
                            gsRow = 6

                            'page down to the next page
                            MoveMe("roll up", 1)
                        Else
                            If Len(Trim(objRx.GetText(gsRow, 3, 8))) > 0 Then
                                gsRow = gsRow + 2
                            Else
                                gsRow = gsRow + 1
                            End If
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            MsgBox("Experienced an exception on ScrapeSpooledFile_Group():  " & ex.ToString)
        End Try
    End Sub

    Sub ScrapeSpooledFile_Member()
        'Lets start scraping!!!! 
        'Dim ssRow As Integer = 4        'SpreadSheet Row
        Dim gsRow As Integer = 6        'Green Screen Row
        Dim bDoneScraping As Boolean = False

        Try
            Do Until bDoneScraping = True

                '  ****  Ran into a situ where "RX CLAIM" got moved over due to the fact that it was placed in the middle of a record  ***************
                ''Added this If statement on 12/22/15
                If gsRow > 24 Then
                    gsRow = 6   'Reset to the beginning

                    'page down to the next page
                    MoveMe("roll up", 1)

                    'loop to the next one
                    Continue Do
                End If
                '  ***********************************************************************************************************************************

                If UCase(Trim(objRx.GetText(gsRow, 90, 14))) = "R X  C L A I M" Then
                    gsRow = gsRow + 5

                    If gsRow > 24 Then

                        gsRow = (gsRow - 25) + 6        'This will reset it back to row 6 and add on up to 4 more rows (for the header)

                        'page down to the next page
                        MoveMe("roll up", 1)

                        'loop to the next one
                        Continue Do
                    End If
                End If

                If LCase(Trim(objRx.GetText(gsRow, 3, 8))) = "carrier:" Then
                    'get the summary info
                    ScrapeSummary_Member(int_TabNum, gsRow)

                    'MsgBox("we are all done (sorta)")
                    bDoneScraping = True
                Else
                    Dim errField1, errField2, errField3, errField4 As String
                    Dim bAppend As Boolean = False

                    'If the length of this is zero then we are appending to what we already have....
                    If Len(Trim(objRx.GetText(gsRow, 3, 10))) < 1 Then
                        bAppend = True
                    End If

                    If bAppend = False Then
                        'Insert as we go...
                        objWorkbook1.Worksheets(int_TabNum).Rows(SSRowNumber).Insert()

                        'Carrier
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsRow, 3, 10))

                        'Account
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsRow, 13, 15))

                        'Group
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 3).Value = Trim(objRx.GetText(gsRow, 29, 15))


                        If gsRow = 24 Then
                            gsRow = 6   'Reset to the beginning

                            'page down to the next page
                            MoveMe("roll up", 1)
                        Else
                            If Len(Trim(objRx.GetText(gsRow, 3, 8))) > 0 Then
                                gsRow = gsRow + 1
                            End If
                        End If


                        'MemberID
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 4).Value = Trim(objRx.GetText(gsRow, 3, 19))

                        'FamilyID
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 5).Value = Trim(objRx.GetText(gsRow, 22, 19))

                        'LastName
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 6).Value = Trim(objRx.GetText(gsRow, 41, 26))

                        'FirstName
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 7).Value = Trim(objRx.GetText(gsRow, 67, 16))

                        'MI
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 8).Value = Trim(objRx.GetText(gsRow, 83, 1))

                        'Sex
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 9).Value = Trim(objRx.GetText(gsRow, 87, 1))

                        'Birthdate
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 10).Value = Trim(objRx.GetText(gsRow, 91, 10))

                        'Error Messages(s)
                        errField1 = Trim(objRx.GetText(gsRow, 105, 28))


                        'Page Right
                        'F20
                        waitForMe()
                        MoveMe("pf20", 1)
                        waitForMe()

                        errField2 = Trim(objRx.GetText(gsRow, 1, 60))

                        'Page Left
                        'f19
                        waitForMe()
                        MoveMe("pf19", 1)
                        waitForMe()


                        If gsRow = 24 Then
                            gsRow = 6   'Reset to the beginning

                            'page down to the next page
                            MoveMe("roll up", 1)
                        Else
                            If Len(Trim(objRx.GetText(gsRow, 3, 8))) > 0 Then
                                gsRow = gsRow + 1
                            End If
                        End If


                        errField3 = Trim(objRx.GetText(gsRow, 105, 28))

                        'Page Right
                        'F20
                        waitForMe()
                        MoveMe("pf20", 1)
                        waitForMe()

                        errField4 = Trim(objRx.GetText(gsRow, 2, 7))

                        'Error Field
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 12).Value = Trim(objRx.GetText(gsRow, 10, 50))

                        'Page Left
                        'f19
                        waitForMe()
                        MoveMe("pf19", 1)
                        waitForMe()

                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber, 11).Value = errField1 & " " & errField2 & " " & errField3 & " " & errField4

                        SSRowNumber = SSRowNumber + 1

                        'If gsRow = 23 Or gsRow = 24 Then
                        If gsRow = 24 Then
                            gsRow = 6   'Reset to the beginning

                            'page down to the next page
                            MoveMe("roll up", 1)
                        Else
                            If Len(Trim(objRx.GetText(gsRow, 3, 8))) > 0 Then
                                gsRow = gsRow + 2
                            Else
                                gsRow = gsRow + 1
                            End If
                        End If

                    Else
                        'This means it is carry-over from the page before...

                        'Group Name
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 4).Value = objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 4).Value & " " & Trim(objRx.GetText(gsRow, 45, 40))

                        'Error Messages
                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 5).Value = objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 5).Value & " " & Trim(objRx.GetText(gsRow, 86, 35))

                        'Error Field
                        errField1 = Trim(objRx.GetText(gsRow, 121, 12))

                        If gsRow + 1 = 25 And Trim(objRx.GetText(gsRow + 1, 121, 12)) = "More..." Then
                            errField3 = Trim(objRx.GetText(gsRow + 1, 121, 12))
                        Else
                            errField3 = ""
                        End If

                        'F20
                        waitForMe()
                        MoveMe("pf20", 1)
                        waitForMe()

                        errField2 = Trim(objRx.GetText(gsRow, 2, 20))

                        'f19
                        waitForMe()
                        MoveMe("pf19", 1)
                        waitForMe()

                        objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 6).Value = objWorkbook1.Worksheets(int_TabNum).Cells(SSRowNumber - 1, 6).Value & " " & errField1 & " " & errField2

                        'If gsRow = 23 Or gsRow = 24 Then
                        If gsRow = 24 Then
                            gsRow = 6

                            'page down to the next page
                            MoveMe("roll up", 1)
                        Else
                            If Len(Trim(objRx.GetText(gsRow, 3, 8))) > 0 Then
                                gsRow = gsRow + 2
                            Else
                                gsRow = gsRow + 1
                            End If
                        End If
                    End If
                End If
            Loop
        Catch ex As Exception
            MsgBox("Experienced an exception on ScrapeSpooledFile_Member():  " & ex.ToString)
        End Try
    End Sub

    Sub ScrapeSummary_Group(spreadsheetTab, GreenScreenRow)
        Dim sCarrier As String
        Dim iGroupRec As Integer
        Dim iGroupErr As Integer
        Dim gsr As Integer = GreenScreenRow
        Dim sSumDetails As String = "Finished the Group Summary for:  "

        Try
            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Carrier
            sSumDetails = sSumDetails & Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Account
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Group
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Identifier
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            sCarrier = Trim(objRx.GetText(gsr, 23, 20))

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = sCarrier
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Groups Input
            gsr = getScrapeSummary_GreenScreenRow(gsr)

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            iGroupRec = Trim(objRx.GetText(gsr, 23, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(iGroupRec)
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Groups Accepted
            gsr = getScrapeSummary_GreenScreenRow(gsr)

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(Trim(objRx.GetText(gsr, 23, 20)))
            SSRowNumber = SSRowNumber + 1
            'spreadsheetRow = spreadsheetRow + 1
            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Groups Rejected
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            iGroupErr = Trim(objRx.GetText(gsr, 23, 20))

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(iGroupErr)
            SSRowNumber = SSRowNumber + 1
            'spreadsheetRow = spreadsheetRow + 1
            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'End of the Report
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(Trim(objRx.GetText(gsr, 23, 20)))
            SSRowNumber = SSRowNumber + 1

            AddCarrierInfo(Me.iCarrierType, iGroupRec, iGroupErr, 0, 0)

            'txtLog.AppendText(sSumDetails)
            'txtLog.Select(txtLog.TextLength, 0)
            'txtLog.ScrollToCaret()

            txtLog.Text = sSumDetails & vbCrLf & txtLog.Text

        Catch ex As Exception
            MsgBox("Experienced an exception on ScrapeSummary_Group():  " & ex.ToString)
        End Try
    End Sub

    Sub ScrapeSummary_Member(spreadsheetTab, GreenScreenRow)
        Dim sCarrier As String
        Dim iMemberRec As Integer
        Dim iMemberErr As Integer
        Dim gsr As Integer = GreenScreenRow
        Dim sSumDetails As String = "Finished the Member Summary for:  "

        Try
            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Carrier
            sSumDetails = sSumDetails & Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Account
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Group
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = Trim(objRx.GetText(gsr, 23, 20)) & "   -   " & Trim(objRx.GetText(gsr, 43, 50))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Identifier
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            sCarrier = Trim(objRx.GetText(gsr, 23, 20))

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = sCarrier
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Groups Input
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            iMemberRec = Trim(objRx.GetText(gsr, 23, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(iMemberRec)
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Groups Accepted
            gsr = getScrapeSummary_GreenScreenRow(gsr)

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(Trim(objRx.GetText(gsr, 23, 20)))
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'Groups Rejected
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            iMemberErr = Trim(objRx.GetText(gsr, 23, 20))

            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(iMemberErr)
            SSRowNumber = SSRowNumber + 1

            objWorkbook1.Worksheets(spreadsheetTab).Rows(SSRowNumber).Insert()

            'End of the Report
            gsr = getScrapeSummary_GreenScreenRow(gsr)
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 1).Value = Trim(objRx.GetText(gsr, 3, 20))
            objWorkbook1.Worksheets(spreadsheetTab).Cells(SSRowNumber, 2).Value = CStr(Trim(objRx.GetText(gsr, 23, 20)))
            SSRowNumber = SSRowNumber + 1

            AddCarrierInfo(Me.iCarrierType, 0, 0, iMemberRec, iMemberErr)

            txtLog.Text = sSumDetails & vbCrLf & txtLog.Text

            'txtLog.AppendText(sSumDetails)
            'txtLog.Select(txtLog.TextLength, 0)
            'txtLog.ScrollToCaret()

        Catch ex As Exception
            MsgBox("Experienced an exception on ScrapeSummary_Member():  " & ex.ToString)
        End Try
    End Sub

    Sub waitOnMe(intHowLong)
        objRx.Wait(intHowLong)
    End Sub

    Sub waitForMe()
        objWait.WaitForAppAvailable()
        System.Threading.Thread.Sleep(10)
        objWait.WaitForInputReady()
    End Sub

    Sub GetUsername()
        Dim objNet      'This will get the username of the person logged into the PC running this Macro

        Try
            objNet = CreateObject("WScript.NetWork")
            usrNm = objNet.UserName
        Catch ex As Exception
            MsgBox("Experienced an exception on GetUsername():  " & ex.ToString)
        End Try
    End Sub

    Sub MoveMe(command, amount)
        'Do what the command says and do it as many times as the amount says
        'Most common commands will be "tab" and "pf12"

        Dim i As Integer

        Try
            For i = 1 To amount
                waitForMe()
                objRx.SendKeys("[" & command & "]")

                'MsgBox("Check here if we have a RED X")

                waitForMe()
            Next
        Catch ex As Exception
            MsgBox("Experienced an exception on MoveMe():  " & ex.ToString)
        End Try
    End Sub

    Sub MoveMe2(command, r, c)
        Try
            waitForMe()
            objRx.SendKeys("[" & command & "]", r, c)
            waitForMe()
        Catch ex As Exception
            MsgBox("Experienced an exception on MoveMe2():  " & ex.ToString)
        End Try
    End Sub

    Sub TypeMe(value)
        Try
            waitForMe()
            'Enter in the value provided
            objRx.SetText(value)
            waitForMe()
        Catch ex As Exception
            MsgBox("Experienced an exception on TypeMe():  " & ex.ToString)
        End Try
    End Sub

    Sub SettingText(text, row, col)
        Try
            waitForMe()
            objRx.SetText(text, row, col)
            waitForMe()
        Catch ex As Exception
            MsgBox("Experienced an exception on SettingText():  " & ex.ToString)
        End Try
    End Sub

    Public Sub IsRightScreenName(scrName, row, col, mil)
        Try
            If (objRx.WaitForString(scrName, row, col, mil, True)) Then    'This will wait up to the Milliseconds provided
                'Do Nothing...because we are on the desired screen
            Else
                MsgBox("stop...we have detected that you are not on the expected screen.  Please look into.  scrName is:  " & scrName & " row is:  " & row & " col is:  " & col & " mil is:  " & mil)
            End If
        Catch ex As Exception
            MsgBox("Experienced an exception on IsRightScreenName():  " & ex.ToString)
        End Try
    End Sub

    'Private Sub BackgroundWorker_BandA_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_BandA.DoWork
    '    Err.Clear()       'This will clear any pre-existing errors

    '    Dim con, rs, sql ', claimsCount
    '    Dim fldCount, iCol

    '    On Error Resume Next

    '    con = CreateObject("ADODB.Connection")

    '    'This is needed to finish out long queries 
    '    'ERROR:  SQL0666 - SQL query exceeds specified time limit or storage limit
    '    'http://www-01.ibm.com/support/docview.wss?uid=nas8N1017615
    '    ' ---> Timeouts occur based on how long the DB2 UDB for iSeries query optimizer estimates a query will run, not the actual execution time. The accuracy of this estimate is directly related to the information available to the optimizer. This information includes statistics on the data acquired through existing indexes. If the proper indexes are not in place, the estimate may be very poor (and the query may not perform well).
    '    con.CommandTimeout = 0
    '    'con.open(ConnectionString)

    '    rs = CreateObject("ADODB.Recordset")







    '    '****************************************************************************************************************************************************************************
    '    'WAS
    '    ''This is the query that will run (it includes the RxClaim#s)
    '    'sql = GetPostQueryString

    '    ''***Excel cell limit is 32,767 characters...so we have to split this query up
    '    'objWorksheet1.Cells(4, 10).Value = sql.ToString.Substring(0, 32700)                             'Left(sql, 32700)                           'Grabbing the left most 32,700 characters
    '    'objWorksheet1.Cells(5, 10).Value = sql.ToString.Substring(32700, ((sql.length - 1) - 32700))    'Mid(sql, 32700, (sql.length - 32700))      'Grabbing char starting at 32,700 for the remainder

    '    ''This will help with how to open a recordset...http://www.w3schools.com/ado/ado_ref_recordset.asp
    '    'rs.open(sql, con, 0, 1, 1)          'Here we are opening the recordset with the query and connection string
    '    '****************************************************************************************************************************************************************************


    '    '' Auto-fit the column widths and row heights
    '    'Dim ObjRange

    '    'ObjRange = objWorksheet5.UsedRange
    '    'ObjRange.EntireColumn.Autofit()
    'End Sub




    'Private Sub BackgroundWorker_BandA_Alt_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker_BandA_Alt.DoWork


    '    Dim con, rs, sql ', claimsCount
    '    Dim fldCount, iCol


    'End Sub



    '*********************************************************************************************************************************************************************
    'Functions

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        'Me.lblClientStatus.Text = "almost..."
    End Sub
End Class

Public Class Carrier
    Public jn As String             'JobName
    Public tp As Integer            '1=Original and 2=Recycle
    Public grpRec As Integer        'Group Received
    Public grpRej As Integer        'Group Rejected
    Public mbrRec As Integer        'Member Received
    Public mbrRej As Integer        'Member Rejected

    Sub New(ByVal JobName As String, ByVal Type As Integer, ByVal GroupRec As Integer, ByVal GroupRej As Integer, ByVal MemberRec As Integer, ByVal MemberRej As Integer)
        jn = JobName
        tp = Type
        grpRec = GroupRec
        grpRej = GroupRej
        mbrRec = MemberRec
        mbrRej = MemberRej
    End Sub
End Class

Public Class Client
    Public nm As String         'Name
    Public elpTime As String    'ElapsedTime

    Sub New(ByVal Name As String, ByVal ElapsedTime As String)
        nm = Name
        elpTime = ElapsedTime
    End Sub
End Class