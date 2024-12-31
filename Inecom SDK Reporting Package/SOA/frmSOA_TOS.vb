Imports System.IO
Imports System.Threading
Imports SAPbobsCOM

Public Class frmSOA_TOS

#Region "Global Variables"
    Private BPCode As String
    Private AsAtDate As DateTime
    Private FromDate As DateTime
    Private IsBBF As String
    Private IsGAT As String
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Private g_sARSOARunningDate As String = ""

    ' IMPORTANT! Choose the correct company before compiling
    Private Const ClientCompany As CompanyCode = CompanyCode.General
    Private Const EmbeddedType As Boolean = False

    Private oFormSOA_TOS As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
#End Region

#Region "Intialize Application"
    Public Sub New()
        Try
            'If Not NotesSetup() Then
            '    MsgBox("Error creating Database '@NCM_SOC2'")
            '    Exit Sub
            'End If
        Catch ex As Exception
            MsgBox("[frmSOA_TOS].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        Dim oItem As SAPbouiCOM.Item

        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_SOA_TOS.srf") Then
            oFormSOA_TOS = SBO_Application.Forms.Item("NCM_SOA_TOS")
            If ClientCompany = CompanyCode.AE Then
                oFormSOA_TOS.Items.Item("ckLogo").Visible = False
            End If
            oFormSOA_TOS.Title = "A/R SOA By Project"
            oFormSOA_TOS.Items.Item("lbStatus").FontSize = 10
            oFormSOA_TOS.Items.Item("lbStyleOpt").TextStyle = 4

            For Each oItem In oFormSOA_TOS.Items
                oItem.Visible = True
            Next

            SetDatasource()
            SetupChooseFromList()
            RetrieveNotes()
            oFormSOA_TOS.Visible = True
        Else
            Try
                If oFormSOA_TOS.Visible = False Then
                    oFormSOA_TOS.Close()
                Else
                    oFormSOA_TOS.Select()
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub
    Private Sub SetDatasource()
        Try
            Dim oRecordset As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            With oFormSOA_TOS.DataSources.UserDataSources
                .Add("DateType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("DateAsAt", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("BPCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                .Add("Period", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("Logo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("HDR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("BBF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("SNP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("GAT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("HAS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("HFN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("EXC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("Notes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1000)
                .Add("cbBPGrp", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10) '2/3/09 Cherine
                .Add("cbBusiness", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30) '9/6/09 Cherine
            End With

            oEdit = oFormSOA_TOS.Items.Item("etBPCode").Specific
            oEdit.DataBind.SetBound(True, "", "BPCode")
            oEdit = oFormSOA_TOS.Items.Item("etDateAsAt").Specific
            oEdit.DataBind.SetBound(True, "", "DateAsAt")
            oEdit.Value = Now.ToString("yyyyMMdd")
            oEdit = oFormSOA_TOS.Items.Item("etNotes").Specific
            oEdit.DataBind.SetBound(True, "", "Notes")

            oCombo = oFormSOA_TOS.Items.Item("cbDateType").Specific
            oCombo.ValidValues.Add("0", "Document Date")
            oCombo.ValidValues.Add("1", "Due Date")
            oCombo.ValidValues.Add("2", "Posting Date")
            oCombo.DataBind.SetBound(True, "", "DateType")
            oFormSOA_TOS.DataSources.UserDataSources.Item("DateType").ValueEx = "0"

            oCombo = oFormSOA_TOS.Items.Item("cbPrdType").Specific
            oCombo.ValidValues.Add("0", "Every 30 Days")
            oCombo.ValidValues.Add("1", "Every Month")
            oCombo.DataBind.SetBound(True, "", "Period")
            oFormSOA_TOS.DataSources.UserDataSources.Item("Period").ValueEx = "0"

            '2/3/09 Cherine
            oCombo = oFormSOA_TOS.Items.Item("cbBPGrp").Specific
            oRecordset.DoQuery("select ""GroupCode"", ""GroupName"" from " & oCompany.CompanyDB & ".""OCRG"" where ""GroupType"" = 'C'")
            Do Until oRecordset.EoF
                oCombo.ValidValues.Add(oRecordset.Fields.Item(0).Value, oRecordset.Fields.Item(1).Value)
                oRecordset.MoveNext()
            Loop
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            '9/6/09 Cherine
            oCombo = oFormSOA_TOS.Items.Item("cbBusiness").Specific
            oCombo.ValidValues.Add("A", "ALL")
            oCombo.ValidValues.Add("S", "Stevedoring")
            oCombo.ValidValues.Add("F", "Forklift")
            oCombo.ValidValues.Add("E", "Exception")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            oCheck = oFormSOA_TOS.Items.Item("ckLogo").Specific
            oCheck.DataBind.SetBound(True, "", "Logo")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            If ClientCompany = CompanyCode.AMS Then
                oFormSOA_TOS.Items.Item("ckLogo").Enabled = False
            End If

            oCheck = oFormSOA_TOS.Items.Item("ckHDR").Specific
            oCheck.DataBind.SetBound(True, "", "HDR")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormSOA_TOS.Items.Item("ckBBF").Specific
            oCheck.DataBind.SetBound(True, "", "BBF")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormSOA_TOS.Items.Item("ckSNP").Specific
            oCheck.DataBind.SetBound(True, "", "SNP")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormSOA_TOS.Items.Item("ckGAT").Specific
            oCheck.DataBind.SetBound(True, "", "GAT")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormSOA_TOS.Items.Item("ckHAS").Specific
            oCheck.DataBind.SetBound(True, "", "HAS")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormSOA_TOS.Items.Item("ckHFN").Specific
            oCheck.DataBind.SetBound(True, "", "HFN")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormSOA_TOS.Items.Item("ckExc").Specific
            oCheck.DataBind.SetBound(True, "", "EXC")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
        Catch ex As Exception
            SBO_Application.MessageBox("[SetDatasource] : " & ex.Message)
        End Try
    End Sub
    Private Sub ShowStatus(ByVal sStatus As String)
        Try
            Dim oStaticText As SAPbouiCOM.StaticText = oFormSOA_TOS.Items.Item("lbStatus").Specific
            oStaticText.Caption = sStatus
        Catch ex As Exception
            SBO_Application.MessageBox("[ShowStatus] : " & ex.Message)
        End Try
    End Sub
    Private Sub LoadViewer()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            Dim iCount As Integer = -1
            oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)) FROM DUMMY")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                g_sARSOARunningDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
            End If

            sQuery = "SELECT COUNT(*) FROM " & oCompany.CompanyDB & ".""@NCM_SOC"" WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
            oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                iCount = oRec.Fields.Item(0).Value
            End If

            If iCount > 0 Then
                Dim frm As New Hydac_FormViewer

                oCombo = oFormSOA_TOS.Items.Item("cbDateType").Specific
                If oCombo.Selected Is Nothing Then
                    oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                frm.Report = oCombo.Selected.Value

                oCombo = oFormSOA_TOS.Items.Item("cbPrdType").Specific
                If oCombo.Selected Is Nothing Then
                    oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If

                frm.IsShared = IsSharedFileExist()  '26/11/2011 - Erwine
                frm.SharedReportName = g_sReportFilename
                frm.ARSOARunningDate = g_sARSOARunningDate & oCompany.UserName.Trim
                frm.Period = oCombo.Selected.Value
                frm.DBUsernameViewer = DBUsername
                frm.DBPasswordViewer = DBPassword
                frm.Username = oCompany.UserName
                frm.AsAtDate = AsAtDate.ToString("yyyyMMdd")

                '9/6/09 Cherine
                oCombo = oFormSOA_TOS.Items.Item("cbBusiness").Specific
                frm.LineBusiness = oCombo.Selected.Description
                oCheck = oFormSOA_TOS.Items.Item("ckLogo").Specific
                frm.HideLogo = IIf(oCheck.Checked, True, False)
                oCheck = oFormSOA_TOS.Items.Item("ckHDR").Specific
                frm.HideHeader = IIf(oCheck.Checked, True, False)
                oCheck = oFormSOA_TOS.Items.Item("ckBBF").Specific
                frm.IsBBF = IIf(oCheck.Checked, 1, 0)
                oCheck = oFormSOA_TOS.Items.Item("ckSNP").Specific
                frm.IsSNP = IIf(oCheck.Checked, 1, 0)
                oCheck = oFormSOA_TOS.Items.Item("ckGAT").Specific
                frm.IsGAT = IIf(oCheck.Checked, 1, 0)
                oCheck = oFormSOA_TOS.Items.Item("ckHAS").Specific
                frm.IsHAS = IIf(oCheck.Checked, 1, 0)
                oCheck = oFormSOA_TOS.Items.Item("ckHFN").Specific
                frm.IsHFN = IIf(oCheck.Checked, 1, 0)
                frm.ReportName = ReportName.SOA_TOS
                frm.CompanySOA = ClientCompany
                frm.ShowDialog()
            Else
                SBO_Application.StatusBar.SetText("No data found.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

        Catch ex As Exception
            SBO_Application.MessageBox("[LoadViewer] : " & ex.Message)
        End Try
    End Sub
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.SOA_TOS)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AR SOA By Project].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition

        Try
            oCFLs = oFormSOA_TOS.ChooseFromLists

            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "2"
            oCFLCreation.UniqueID = "CFL_BPCode"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oFormSOA_TOS.Items.Item("etBPCode").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BPCode"
            oEditLn.ChooseFromListAlias = "CardCode"
 
        Catch ex As Exception
            Throw New Exception("[ARSOA].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub

#End Region

#Region "Logic Function"
    Private Function NotesSetup() As Boolean
        Dim bSuccess As Boolean = False
        Dim sQuery As String = ""
        Dim sCurrSchema As String = ""
        Dim iCount As Integer = 0
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

        Try
            sQuery = " SELECT current_schema FROM DUMMY "
            oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                sCurrSchema = oRec.Fields.Item(0).Value
            End If

            If sCurrSchema.Trim <> "" Then
                sQuery = "  select Count(*) from sys.objects "
                sQuery &= " where ""SCHEMA_NAME"" = '" & sCurrSchema & "' "
                sQuery &= " AND ""OBJECT_TYPE"" = 'TABLE '"
                sQuery &= " AND ""OBJECT_NAME"" = '@NCM_SOC2' "
                oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    iCount = oRec.Fields.Item(0).Value
                End If

                If iCount <= 0 Then
                    sQuery = " CREATE TABLE ""@NCM_SOC2"" "
                    sQuery &= " (ID         NVARCHAR(8)         NOT NULL,"
                    sQuery &= " Notes      NVARCHAR(2000)      NOT NULL,"
                    sQuery &= " Image    BLOB)"
                    oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)

                    sQuery = " INSERT INTO ""@NCM_SOC2"" "
                    sQuery &= " VALUES ("
                    sQuery &= " '1',"
                    sQuery &= " 'Note:   Any payments received after end of the month will be shown in next month''s statement."
                    sQuery &= "          If you do not agree with the above statement, please inform us immediately.'"
                    sQuery &= " , NULL) "
                    oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                Else
                    iCount = 0
                    sQuery = " Select Count(*) from ""@NCM_SOC2"" WHERE ""ID"" = '1' "
                    oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If oRec.RecordCount > 0 Then
                        iCount = Convert.ToInt32(oRec.Fields.Item(0).Value)
                    End If

                    If iCount <= 0 Then
                        sQuery = " INSERT INTO ""@NCM_SOC2"" "
                        sQuery &= " VALUES ("
                        sQuery &= " '1',"
                        sQuery &= " 'Note:   Any payments received after end of the month will be shown in next month''s statement."
                        sQuery &= "          If you do not agree with the above statement, please inform us immediately.'"
                        sQuery &= " , NULL) "
                        oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRec.DoQuery(sQuery)
                    End If
                End If
            End If

            oRec = Nothing
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[NotesSetup] : " & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Private Sub RetrieveNotes()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""NOTES"" FROM ""@NCM_SOC2"" WHERE ""ID"" ='1'")
            If oRec.RecordCount > 0 Then
                oFormSOA_TOS.DataSources.UserDataSources.Item("Notes").ValueEx = oRec.Fields.Item(0).Value
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOAProject].[RetrieveNotes]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function SaveSettings() As Boolean
        Dim Notes As String = ""
        Dim BitmapPath As String = ""
        Dim ImagePath As String = ""
        Dim sQuery As String
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

        Try
            ShowStatus("Status: Saving Settings...")
            Notes = oFormSOA_TOS.DataSources.UserDataSources.Item("Notes").ValueEx
            Notes = Notes.Replace("'", "''")
            sQuery = "SELECT IFNULL(""BitmapPath"",'') FROM " & oCompany.CompanyDB & ".""OADP"" "
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                BitmapPath = oRec.Fields.Item(0).Value
            End If

            If ClientCompany <> CompanyCode.AMS Then
                ImagePath = BitmapPath & oCompany.CompanyDB & ".bmp"
                If File.Exists(ImagePath) = False Then
                    ImagePath = BitmapPath & oCompany.CompanyDB & ".jpg"
                    If File.Exists(ImagePath) = False Then
                        ImagePath = BitmapPath & oCompany.CompanyDB & ".png"
                        If File.Exists(ImagePath) = False Then
                            ImagePath = BitmapPath & oCompany.CompanyDB & ".tiff"
                            If File.Exists(ImagePath) = False Then
                                ImagePath = ""
                            End If
                        End If
                    End If
                End If
            End If

            sQuery = "UPDATE " & oCompany.CompanyDB & ".""@NCM_SOC2"" SET ""NOTES"" ='" & Notes & "' WHERE ""ID"" = '1'"
            oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOAProject].[SaveImages] : " & ex.Message)
            Return False
        End Try
    End Function
    Private Function MainCtrlNONEMBEDDED() As Boolean
        Dim sDate As String = String.Empty
        Dim sBBF As String = "N"
        Dim bSuccess As Boolean = False
        Dim iRowsAffected As Integer = 0
        Dim sQuery As String = String.Empty

        Try
            'Get BPCode
            oEdit = oFormSOA_TOS.Items.Item("etBPCode").Specific
            BPCode = CType(IIf(oEdit.Value = "", "%", oEdit.Value.Replace("*", "%")), String).Trim
            BPCode = oEdit.Value

            'Get AsAtDate, FromDate
            oEdit = oFormSOA_TOS.Items.Item("etDateAsAt").Specific
            sDate = oEdit.Value.Trim
            If sDate = "" Then Throw New Exception("Error: As At Date is empty!")
            AsAtDate = New DateTime(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
            FromDate = New DateTime(Left(sDate, 4), Mid(sDate, 5, 2), "01")

            'Get IsBBF
            oCheck = oFormSOA_TOS.Items.Item("ckBBF").Specific
            If oCheck.Checked Then IsBBF = "Y" Else IsBBF = "N"

            'Get IsGAT
            oCheck = oFormSOA_TOS.Items.Item("ckGAT").Specific
            If oCheck.Checked Then IsGAT = "Y" Else IsGAT = "N"

            '2/3/09 Cherine - GetBPGroup
            oCombo = oFormSOA_TOS.Items.Item("cbBPGrp").Specific

            'Set the query
            sQuery = " CALL SP_SOA_TOS ('"
            sQuery &= g_sARSOARunningDate & oCompany.UserName & "','','','','','','',"
            sQuery &= "'" & BPCode & "',"
            sQuery &= "'" & FromDate.ToString("yyyyMMdd") & "',"
            sQuery &= "'" & AsAtDate.ToString("yyyyMMdd") & "',"
            sQuery &= "'" & IsBBF & "',"
            sQuery &= "'" & IsGAT & "',"
            sQuery &= "'" & oCombo.Selected.Value & "'," '2/3/09 Cherine

            oCheck = oFormSOA_TOS.Items.Item("ckExc").Specific
            If oCheck.Checked Then sQuery &= "'1') " Else sQuery &= "'0') "

            'CREATE PROCEDURE SP_SOA_TOS
            '(@UserName		NVARCHAR(20),	-- Capture the User running this Procedure
            ' @BPCODEFR 		NVARCHAR(20),
            ' @BPCODETO		NVARCHAR(20),
            ' @BPGPFR		NVARCHAR(20),
            ' @BPGPTO		NVARCHAR(20),
            ' @SLPNAMEFR		NVARCHAR(32),
            ' @SLPNAMETO		NVARCHAR(32),
            ' @BPCode		NVARCHAR(20),	-- '%' for all BP or individual BPCode
            ' @PeriodFr		DATETIME,		-- 1st Day of the As At Date's Month
            ' @PeriodTo		DATETIME,		-- SOA As At Date
            ' @BalFwd		NVARCHAR(1),	-- 'Y' to print Balance B/F line
            ' @SortOrdr		NVARCHAR(1),	-- Indicate the Printing Order of SOA
            ' @BPGroup		INT,			-- 2/3/09 Cherine - Whether is internal/external customer
            ' @Option		INT) 	

            SBO_Application.StatusBar.SetText("Completed Successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOAProject].[MainCtrlNONEMBEDDED]" & vbNewLine & ex.Message)
        End Try
        Return bSuccess
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.ItemUID
                    Case "ckHFN"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            oFormSOA_TOS.Items.Item("etBPCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "btnExecute"
                                If SaveSettings() Then
                                    If Not EmbeddedType Then
                                        If MainCtrlNONEMBEDDED() Then
                                            Dim myThread As Thread = New Thread(New ThreadStart(AddressOf LoadViewer))
                                            myThread.SetApartmentState(ApartmentState.STA)
                                            myThread.Start()
                                        End If
                                    End If
                                End If
                            Case "ckHFN"
                                oFormSOA_TOS.Items.Item("etNotes").Enabled = Not (oFormSOA_TOS.Items.Item("etNotes").Enabled)
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvent = pVal
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvent.SelectedObjects
                        If (Not oDataTable Is Nothing) Then
                            Dim sTemp As String = ""
                            With oFormSOA_TOS.DataSources.UserDataSources
                                Select Case oCFLEvent.ChooseFromListUID
                                    Case "CFL_BPCode"
                                        sTemp = oDataTable.GetValue("CardCode", 0)
                                        .Item("BPCode").ValueEx = sTemp
                                        Exit Select
                                    Case Else
                                        Exit Select
                                End Select
                            End With
                            Return True
                        End If
                End Select
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOAProject].[ItemEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class