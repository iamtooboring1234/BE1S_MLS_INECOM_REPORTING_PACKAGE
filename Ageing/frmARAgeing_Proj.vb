Imports System.IO
Imports System.Threading
Imports System.Xml
Imports System.Data.Common

Public Class frmARAgeing_Proj

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oItem As SAPbouiCOM.Item
    Private oStatic As SAPbouiCOM.StaticText
    Private oPictureBox As SAPbouiCOM.PictureBox
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Private oCheck As SAPbouiCOM.CheckBox

    Private sTxtFormat As String = "txtB{0}txt"
    Private sValFormat As String = "txtB{0}Val"
    Private sFTxtFormat As String = "U_Bucket{0}Txt"
    Private sFValFormat As String = "U_Bucket{0}Val"

    Private iCount As Integer = 1
    Private sTxtBTxt As String() = New String(10) {}
    Private sTxtBVal As Integer() = New Integer(10) {}
    Private sExcelPath As String = String.Empty
    Private bIsSaveRunning As Boolean = True
    Private bIsCancel As Boolean = False
    Private bIsExportToExcel As Boolean = False
    Private sRptType As String = String.Empty
    Private g_sProjects As String = ""
    Private g_sProjFr As String = ""
    Private g_sProjTo As String = ""
    Private g_sARAGERunningDate As String = ""

    Public HANADbConnection As DbConnection
    Public _DbProviderFactoryObject As DbProviderFactory
    Public ProviderName As String = "System.Data.Odbc"

#End Region

#Region "Initialization"
    Public Sub ShowForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmARAgeing_Proj.srf") Then
            oForm = SBO_Application.Forms.Item("ncmARAgeing_Proj")
            oForm.Title = "AR Ageing with Project Report"
            oPictureBox = oForm.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"

            AddDataSource()
            PopulateData("Y")
            SetupChooseFromList()
            oForm.Visible = True
        Else
            Try
                oForm = SBO_Application.Forms.Item("ncmARAgeing_Proj")
                If oForm.Visible = False Then
                    oForm.Close()
                Else
                    oForm.Select()
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub AddDataSource()
        Try
            Dim sTemp As String = String.Empty

            With oForm.DataSources.UserDataSources
                .Add("BPCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                .Add("dDate", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("AgeBy", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("RptType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("txtBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtBPGFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtBPGTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtSlsFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtSlsTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("chkPage", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("chkExcel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("ckFin", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("tbProjFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("tbProjTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)

                oEdit = oForm.Items.Item("tbProjFr").Specific
                oEdit.DataBind.SetBound(True, "", "tbProjFr")
                oEdit = oForm.Items.Item("tbProjTo").Specific
                oEdit.DataBind.SetBound(True, "", "tbProjTo")

                oEdit = oForm.Items.Item("txtBPCode").Specific
                oEdit.DataBind.SetBound(True, "", "BPCode")
                oEdit = oForm.Items.Item("txtDate").Specific
                oEdit.DataBind.SetBound(True, "", "dDate")
                .Item("dDate").ValueEx = Now.ToString("yyyyMMdd")

                oCombo = oForm.Items.Item("cboAgeBy").Specific
                oCombo.ValidValues.Add("0", "Document Date")
                oCombo.ValidValues.Add("1", "Due Date")
                oCombo.ValidValues.Add("2", "Posting Date")
                oCombo.DataBind.SetBound(True, "", "AgeBy")
                .Item("AgeBy").ValueEx = "0"

                oCombo = oForm.Items.Item("cboRptType").Specific
                oCombo.ValidValues.Add("0", "Details")
                oCombo.ValidValues.Add("1", "Summary")
                oCombo.DataBind.SetBound(True, "", "RptType")
                .Item("RptType").ValueEx = "0"

                For iCount = 1 To 9
                    .Add(String.Format(sTxtFormat, iCount), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                    .Add(String.Format(sValFormat, iCount), SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
                Next

                For iCount = 1 To 5
                    sTemp = String.Format(sTxtFormat, iCount)
                    oEdit = oForm.Items.Item(sTemp).Specific
                    oEdit.DataBind.SetBound(True, String.Empty, sTemp)

                    sTemp = String.Format(sValFormat, iCount)
                    oEdit = oForm.Items.Item(sTemp).Specific
                    oEdit.DataBind.SetBound(True, String.Empty, sTemp)
                Next
                oEdit = oForm.Items.Item("txtBPFr").Specific
                oEdit.DataBind.SetBound(True, "", "txtBPFr")
                oEdit = oForm.Items.Item("txtBPTo").Specific
                oEdit.DataBind.SetBound(True, "", "txtBPTo")
                oEdit = oForm.Items.Item("txtBPGFr").Specific
                oEdit.DataBind.SetBound(True, "", "txtBPGFr")
                oEdit = oForm.Items.Item("txtBPGTo").Specific
                oEdit.DataBind.SetBound(True, "", "txtBPGTo")
                oEdit = oForm.Items.Item("txtSlsFr").Specific
                oEdit.DataBind.SetBound(True, "", "txtSlsFr")
                oEdit = oForm.Items.Item("txtSlsTo").Specific
                oEdit.DataBind.SetBound(True, "", "txtSlsTo")

                oCheck = DirectCast(oForm.Items.Item("chkPage").Specific, SAPbouiCOM.CheckBox)
                oCheck.DataBind.SetBound(True, String.Empty, "chkPage")
                oCheck.ValOff = "0"
                oCheck.ValOn = "1"
                oCheck = oForm.Items.Item("ckFin").Specific
                oCheck.DataBind.SetBound(True, String.Empty, "ckFin")
                oCheck.ValOff = "N"
                oCheck.ValOn = "Y"
                oCheck = oForm.Items.Item("chkExcel").Specific
                oCheck.DataBind.SetBound(True, String.Empty, "chkExcel")
                oCheck.ValOff = "N"
                oCheck.ValOn = "Y"

                oForm.DataSources.UserDataSources.Item("ckFin").ValueEx = "N"

            End With
        Catch ex As Exception
            SBO_Application.MessageBox("[frmARAgeing].[AddDataSource]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub ShowStatus(ByVal sStatus As String)
        Try
            Dim oStaticText As SAPbouiCOM.StaticText = oForm.Items.Item("lbStatus").Specific
            oStaticText.Caption = sStatus
        Catch ex As Exception
            SBO_Application.MessageBox("[frmARAgeing].[ShowStatus]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition

        Try
            oCFLs = oForm.ChooseFromLists
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 2
            oCFLCreation.UniqueID = "cflBPFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oForm.Items.Item("txtBPFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflBPFr"
            oEditLn.ChooseFromListAlias = "CardCode"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "2"
            oCFLCreation.UniqueID = "cflBPTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oForm.Items.Item("txtBPTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflBPTo"
            oEditLn.ChooseFromListAlias = "CardCode"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 10
            oCFLCreation.UniqueID = "CFL_BGFrom"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "GroupType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oForm.Items.Item("txtBPGFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BGFrom"
            oEditLn.ChooseFromListAlias = "GroupCode"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 10
            oCFLCreation.UniqueID = "CFL_BGTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "GroupType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oForm.Items.Item("txtBPGTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BGTo"
            oEditLn.ChooseFromListAlias = "GroupCode"
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

            oEditLn = DirectCast(oForm.Items.Item("txtBPCode").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BPCode"
            oEditLn.ChooseFromListAlias = "CardCode"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 53
            oCFLCreation.UniqueID = "CFL_SPFrom"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtSlsFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_SPFrom"
            oEditLn.ChooseFromListAlias = "SlpName"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 53
            oCFLCreation.UniqueID = "CFL_SPTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtSlsTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_SPTo"
            oEditLn.ChooseFromListAlias = "SlpName"
            ' -----------------------------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ProjectCodes
            oCFLCreation.UniqueID = "CFL_ProjFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCFLCreation.UniqueID = "CFL_ProjTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("tbProjFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_ProjFr"
            oEditLn.ChooseFromListAlias = "PrjCode"

            oEditLn = DirectCast(oForm.Items.Item("tbProjTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_ProjTo"
            oEditLn.ChooseFromListAlias = "PrjCode"

        Catch ex As Exception
            Throw New Exception("[ARAgeing].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub PopulateData(Optional ByVal sVal As String = "")
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sTemp1 As String = String.Empty
            Dim sCheck As String = ""

            If sVal <> "" Then
                sCheck = sVal
            Else
                sCheck = oForm.DataSources.UserDataSources.Item("ckFin").ValueEx
            End If
            Select Case sCheck
                Case "Y"
                    sTxtBTxt(0) = "0-30"
                    sTxtBTxt(1) = "31-60"
                    sTxtBTxt(2) = "61-90"
                    sTxtBTxt(3) = "91-120"
                    sTxtBTxt(4) = ">120"

                    sTxtBVal(0) = 30
                    sTxtBVal(1) = 60
                    sTxtBVal(2) = 90
                    sTxtBVal(3) = 120
                    sTxtBVal(4) = 120

                    oRecord.DoQuery(" SELECT *  FROM " & oCompany.CompanyDB & ".""@NCM_BUCKET"" WHERE ""U_Type"" = 'NCM_AR_AGEING'")
                    If oRecord.RecordCount > 0 Then
                        iCount = 0
                        For iCount = 0 To 4
                            sTemp1 = String.Format(sFTxtFormat, iCount + 1)
                            sTxtBTxt(iCount) = oRecord.Fields.Item(sTemp1).Value

                            sTemp1 = String.Format(sFValFormat, iCount + 1)
                            sTxtBVal(iCount) = oRecord.Fields.Item(sTemp1).Value
                        Next
                    End If
                    For iCount = 0 To 4
                        sTemp1 = String.Format(sTxtFormat, iCount + 1)
                        oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBTxt(iCount)

                        sTemp1 = String.Format(sValFormat, iCount + 1)
                        oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBVal(iCount)
                    Next

                Case "N"
                    Dim oCheck As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    Dim squery As String = ""
                    Dim sAsAtDate As String = "'" & oForm.DataSources.UserDataSources.Item("dDate").ValueEx & "'"

                    'column title - current month
                    squery = "SELECT ""Code"" FROM " & oCompany.CompanyDB & ".""OFPR"" WHERE " & sAsAtDate & " BETWEEN ""F_RefDate"" AND ""T_RefDate"""
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBTxt(0) = oRecord.Fields.Item(0).Value
                    End If

                    'column title - last 1 month
                    'squery = "SELECT code FROM OFPR WHERE dateadd(month,-1," & sAsAtDate & ") between F_RefDate and T_RefDate"
                    squery = "SELECT ""Code"" FROM " & oCompany.CompanyDB & ".""OFPR"" WHERE ADD_MONTHS(" & sAsAtDate & ", -1) BETWEEN ""F_RefDate"" AND ""T_RefDate"""
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBTxt(1) = oRecord.Fields.Item(0).Value
                    End If

                    'column title - last 2 month
                    'squery = "SELECT code FROM OFPR WHERE dateadd(month,-2," & sAsAtDate & ") between F_RefDate and T_RefDate"
                    squery = "SELECT ""Code"" FROM " & oCompany.CompanyDB & ".""OFPR"" WHERE ADD_MONTHS(" & sAsAtDate & ", -2) BETWEEN ""F_RefDate"" AND ""T_RefDate"""
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBTxt(2) = oRecord.Fields.Item(0).Value
                    End If

                    'column title - last 3 month
                    'squery = "SELECT code FROM OFPR WHERE dateadd(month,-3," & sAsAtDate & ") between F_RefDate and T_RefDate"
                    squery = "SELECT ""Code"" FROM " & oCompany.CompanyDB & ".""OFPR"" WHERE ADD_MONTHS(" & sAsAtDate & ", -3) BETWEEN ""F_RefDate"" AND ""T_RefDate"""
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBTxt(3) = oRecord.Fields.Item(0).Value
                    End If

                    'column title - last 4 month
                    sTxtBTxt(4) = "> " & sTxtBTxt(3)

                    'no of days - current month
                    'squery = "select DAY(" & sAsAtDate & ")"
                    squery = "SELECT DAYOFMONTH(" & sAsAtDate & ") FROM DUMMY"

                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBVal(0) = Convert.ToInt32(oRecord.Fields.Item(0).Value)
                    End If

                    'no of days - last 1 month
                    'squery = "select datepart(day, dateadd(s,-1,dateadd(mm, datediff(m,0," & sAsAtDate & "),0)))"
                    squery = " SELECT DAYOFMONTH(Last_day(ADD_MONTHS(" & sAsAtDate & ", -1)))  FROM DUMMY "
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBVal(1) = Convert.ToInt32(sTxtBVal(0)) + Convert.ToInt32(oRecord.Fields.Item(0).Value)
                    End If

                    'no of days - last 2 month
                    'squery = "select datepart(day, dateadd(s,-1,dateadd(mm, datediff(m,0," & sAsAtDate & ")-1,0)))"
                    squery = " SELECT DAYOFMONTH(Last_day(ADD_MONTHS(" & sAsAtDate & ", -2)))  FROM DUMMY "
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBVal(2) = Convert.ToInt32(sTxtBVal(1)) + Convert.ToInt32(oRecord.Fields.Item(0).Value)
                    End If

                    'no of days - last 3 month
                    'squery = "select datepart(day, dateadd(s,-1,dateadd(mm, datediff(m,0," & sAsAtDate & ")-2,0)))"
                    squery = " SELECT DAYOFMONTH(Last_day(ADD_MONTHS(" & sAsAtDate & ", -3)))  FROM DUMMY "
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(squery)
                    If oRecord.RecordCount > 0 Then
                        sTxtBVal(3) = Convert.ToInt32(sTxtBVal(2)) + Convert.ToInt32(oRecord.Fields.Item(0).Value)
                    End If

                    sTxtBVal(4) = Convert.ToInt32(sTxtBVal(3))
                    For iCount = 0 To 4
                        sTemp1 = String.Format(sTxtFormat, iCount + 1)
                        oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBTxt(iCount)

                        sTemp1 = String.Format(sValFormat, iCount + 1)
                        oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBVal(iCount)
                    Next

            End Select
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PopulateData] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub OpenSaveFileDialog()
        Try
            Dim frmSave As New frmSaveDialog
            frmSave.TopMost = True
            frmSave.Show()
            System.Threading.Thread.CurrentThread.Sleep(10)
            frmSave.Activate()

            If (frmSave.SaveFileDialog1.ShowDialog() = DialogResult.OK) Then
                sExcelPath = frmSave.SaveFileDialog1.FileName
            Else
                bIsCancel = True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            bIsSaveRunning = False
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Function IsSharedFileExist() As Boolean
        Try
            '' 1 - File not found thus use local - 0 - File's found thus use shared file - [-1] - error
            Dim RptType As String = String.Empty
            oCombo = oForm.Items.Item("cboRptType").Specific
            RptType = oCombo.Selected.Value
            g_sReportFilename = ""

            Select Case RptType
                Case 0
                    g_sReportFilename = GetSharedFilePath(ReportName.ARAging_Details_Proj)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
                Case 1
                    g_sReportFilename = GetSharedFilePath(ReportName.ARAging_Summary_Proj)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
            End Select
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[A/R AGEING].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub GetProjects()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sProjFr As String = ""
            Dim sProjTo As String = ""
            Dim sQuery As String = ""

            g_sProjects = ""
            g_sProjFr = ""
            g_sProjTo = ""

            oEdit = oForm.Items.Item("tbProjFr").Specific
            sProjFr = oEdit.Value.Trim
            oEdit = oForm.Items.Item("tbProjTo").Specific
            sProjTo = oEdit.Value.Trim

            If sProjFr = "" And sProjTo = "" Then
                g_sProjects = ""
                ' sQuery = "SELECT TOP 1 PrjCode FROM OPRJ ORDER BY PrjCode "
                sQuery = "SELECT TOP 1 ""PrjCode"" FROM " & oCompany.CompanyDB & ".""OPRJ"" ORDER BY ""PrjCode"" "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    g_sProjFr = oRec.Fields.Item(0).Value.ToString.Trim
                End If

                'sQuery = "SELECT TOP 1 PrjCode FROM OPRJ ORDER BY PrjCode DESC"
                sQuery = "SELECT TOP 1 ""PrjCode"" FROM " & oCompany.CompanyDB & ".""OPRJ"" ORDER BY ""PrjCode"" DESC "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    g_sProjTo = oRec.Fields.Item(0).Value.ToString.Trim
                End If
            Else
                g_sProjFr = sProjFr
                g_sProjTo = sProjTo
                g_sProjects = ""

                Select Case sProjFr
                    Case ""
                        If sProjTo <> "" Then
                            'sQuery = " SELECT PrjCode, PrjName FROM [OPRJ] WHERE PrjCode <= '" & sProjTo & "' "
                            sQuery = " SELECT ""PrjCode"", ""PrjName"" FROM " & oCompany.CompanyDB & ".""OPRJ"" WHERE ""PrjCode"" <= '" & sProjTo & "' "
                        End If
                    Case Else
                        If sProjTo <> "" Then
                            'sQuery = " SELECT PrjCode, PrjName FROM [OPRJ] WHERE PrjCode BETWEEN '" & sProjFr & "' AND '" & sProjTo & "'"
                            sQuery = " SELECT ""PrjCode"", ""PrjName"" FROM " & oCompany.CompanyDB & ".""OPRJ"" WHERE ""PrjCode"" BETWEEN '" & sProjFr & "' AND '" & sProjTo & "'"
                        Else
                            'sQuery = " SELECT PrjCode, PrjName FROM [OPRJ] WHERE PrjCode >= '" & sProjTo & "' "
                            sQuery = " SELECT ""PrjCode"", ""PrjName"" FROM " & oCompany.CompanyDB & ".""OPRJ"" WHERE ""PrjCode"" >= '" & sProjTo & "' "
                        End If
                End Select
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    While Not oRec.EoF
                        g_sProjects &= "'" & oRec.Fields.Item(0).Value.ToString.Trim & "',"
                        oRec.MoveNext()
                    End While
                    g_sProjects = g_sProjects.Remove(g_sProjects.Length - 1, 1)
                End If
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[A/R AGEING].[GetProjects] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadViewer()
        oForm.Items.Item("btnExecute").Enabled = False
        Try
            Dim frm As New Hydac_FormViewer
            Dim bIsContinue As Boolean = False
            Dim sBPCodeFr As String = ""
            Dim sBPCodeTo As String = ""
            Dim sBPGrpFr As String = ""
            Dim sBPGrpTo As String = ""
            Dim sSlsFr As String = ""
            Dim sSlsTo As String = ""
            Dim sLocalCurr As String = ""
            Dim iPageBreak As Integer = 0
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oUpdate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            g_sARAGERunningDate = ""
            Try
                Dim oTest As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)) FROM DUMMY")
                If oTest.RecordCount > 0 Then
                    oTest.MoveFirst()
                    g_sARAGERunningDate = Convert.ToString(oTest.Fields.Item(0).Value)
                End If
                oTest = Nothing

                GetProjects()

                If (ExecuteProcedure()) Then
                    'oUpdate.DoQuery("UPDATE [@NCM_AR_AGEING] SET Project = '<Blank>' WHERE ISNULL(Project,'') = ''")
                    oUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oUpdate.DoQuery("UPDATE " & oCompany.CompanyDB & ".""@NCM_AR_AGEING"" SET ""PROJECT"" = '<Blank>' WHERE IFNULL(""PROJECT"", '') = '' AND ""USERNAME"" ='" & g_sARAGERunningDate & oCompany.UserName.Trim & "' ")

                    oUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oUpdate.DoQuery("UPDATE " & oCompany.CompanyDB & ".""@NCM_AR_AGEING"" SET ""NUMATCARD"" = '' WHERE IFNULL(""NUMATCARD"", '') = '' AND ""USERNAME"" ='" & g_sARAGERunningDate & oCompany.UserName.Trim & "' ")

                    oUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oUpdate.DoQuery("UPDATE " & oCompany.CompanyDB & ".""@NCM_AR_AGEING"" SET ""REF2"" = '' WHERE IFNULL(""REF2"", '') = '' AND ""USERNAME"" ='" & g_sARAGERunningDate & oCompany.UserName.Trim & "' ")

                    oUpdate = Nothing

                    bIsContinue = True
                    oEdit = oForm.Items.Item("txtDate").Specific
                    Dim AsAtDate As String = oEdit.Value
                    oCombo = oForm.Items.Item("cboAgeBy").Specific
                    Dim AgeingBy As String = oCombo.Selected.Value
                    Dim sAgeingBy As String = oCombo.Selected.Description

                    oCombo = oForm.Items.Item("cboRptType").Specific
                    Dim RptType As String = oCombo.Selected.Value

                    'Get Parameter Value
                    oEdit = oForm.Items.Item("txtBPFr").Specific
                    sBPCodeFr = oEdit.Value
                    oEdit = oForm.Items.Item("txtBPTo").Specific
                    sBPCodeTo = oEdit.Value
                    oEdit = oForm.Items.Item("txtBPGFr").Specific
                    sBPGrpFr = oEdit.Value
                    oEdit = oForm.Items.Item("txtBPGTo").Specific
                    sBPGrpTo = oEdit.Value
                    oEdit = oForm.Items.Item("txtSlsFr").Specific
                    sSlsFr = oEdit.Value
                    oEdit = oForm.Items.Item("txtSlsTo").Specific
                    sSlsTo = oEdit.Value

                    sLocalCurr = GetLocalCurrency()
                    oEdit = oForm.Items.Item("txtBPCode").Specific
                    Dim sBPCode As String = oEdit.Value
                    oCheck = DirectCast(oForm.Items.Item("chkPage").Specific, SAPbouiCOM.CheckBox)
                    If (oCheck.Checked) Then
                        iPageBreak = 1
                    End If

                    SBO_Application.StatusBar.SetText("Opening Report...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    frm.IsShared = g_bIsShared
                    frm.SharedReportName = g_sReportFilename
                    frm.ReportName = ReportName.ARAging_Details_Proj
                    frm.ARAGERunningDate = g_sARAGERunningDate & oCompany.UserName.Trim
                    frm.DBPasswordViewer = DBPassword
                    frm.DBUsernameViewer = DBUsername
                    frm.Username = oCompany.UserName
                    frm.AsAtDate = AsAtDate
                    frm.AgeBy = AgeingBy
                    frm.BPCode = sBPCode
                    frm.FromProj = g_sProjFr
                    frm.ToProj = g_sProjTo
                    frm.Projects = g_sProjects
                    frm.BPCodeFr = sBPCodeFr
                    frm.BPCodeTo = sBPCodeTo
                    frm.BPGroupFr = sBPGrpFr
                    frm.BPGroupTo = sBPGrpTo
                    frm.SalesEmployeeFr = sSlsFr
                    frm.SalesEmployeeTo = sSlsTo
                    frm.AgingBy = sAgeingBy
                    frm.LocalCurrency = sLocalCurr
                    frm.SectionPageBreak = iPageBreak
                    frm.BucketText = sTxtBTxt
                    frm.BucketValue = sTxtBVal
                    frm.IsExcel = bIsExportToExcel

                    If RptType = 0 Then
                        frm.Text = "AR Ageing Details Report with Project "
                        frm.ReportType = AgeingType.ARAgeing
                    ElseIf RptType = 1 Then
                        frm.Text = "AR Ageing Summary Report with Project "
                        frm.ReportType = AgeingType.ARAgeingSummary
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item("btnExecute").Enabled = True
            End Try
            If bIsContinue Then
                If (bIsExportToExcel) Then
                    'Export To Excel
                    '--------------------------------------------------------------------------------
                    frm.OpenAgingReport()
                Else
                    'Not Export To Excel
                    '--------------------------------------------------------------------------------
                    frm.ShowDialog()
                End If
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[frmARAgeing].[LoadViewer]:" & ex.ToString)
        Finally
        End Try
    End Sub
    Private Function GetLocalCurrency() As String
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""MainCurncy"" FROM " & oCompany.CompanyDB & ".""OADM"" ")
            If (oRec.RecordCount > 0) Then
                Return oRec.Fields.Item(0).Value.ToString()
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[frmARAging].[GetLocalCurrency] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return String.Empty
    End Function
    Private Function ExecuteProcedure() As Boolean
        g_bIsShared = IsSharedFileExist()
        If (g_bIsShared) Then
            If (Not File.Exists(g_sReportFilename)) Then
                SBO_Application.StatusBar.SetText("[AR_AGEING_PROJ]: Crystal Report file is not found in location (" & g_sReportFilename & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If

        Dim sBPCode As String = ""
        Dim sAsAtDate As String = ""
        Dim sBPCodeFr As String = String.Empty
        Dim sBPCodeTo As String = String.Empty
        Dim sBPGrpFr As String = String.Empty
        Dim sBPGrpTo As String = String.Empty
        Dim sSlsFr As String = String.Empty
        Dim sSlsTo As String = String.Empty
        Dim sQuery As String = ""
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            'Get Parameter Value
            oEdit = oForm.Items.Item("txtBPFr").Specific
            sBPCodeFr = oEdit.Value
            oEdit = oForm.Items.Item("txtBPTo").Specific
            sBPCodeTo = oEdit.Value
            oEdit = oForm.Items.Item("txtBPGFr").Specific
            sBPGrpFr = oEdit.Value
            oEdit = oForm.Items.Item("txtBPGTo").Specific
            sBPGrpTo = oEdit.Value
            oEdit = oForm.Items.Item("txtSlsFr").Specific
            sSlsFr = oEdit.Value
            oEdit = oForm.Items.Item("txtSlsTo").Specific
            sSlsTo = oEdit.Value
            oEdit = oForm.Items.Item("txtBPCode").Specific
            ' BPCode = IIf(oEdit.Value.Trim = "", "%", "%" & oEdit.Value.Trim.Replace("*", "%") & "%")
            sBPCode = oEdit.Value
            oEdit = oForm.Items.Item("txtDate").Specific
            sAsAtDate = oEdit.Value
            ' dtAsAtDate = New DateTime(Left(AsAtDate, 4), Mid(AsAtDate, 5, 2), Right(AsAtDate, 2))
            
            oStatic = oForm.Items.Item("lbStatus").Specific
            oStatic.Caption = "Executing Store Procedure. Please wait..."

            sQuery = " CALL SP_AR_AGEING ("
            sQuery &= "'" & g_sARAGERunningDate & oCompany.UserName.Trim & "',"
            sQuery &= "'" & sBPCodeFr.Replace("'", "''") & "',"
            sQuery &= "'" & sBPCodeTo.Replace("'", "''") & "',"
            sQuery &= "'" & sBPGrpFr & "',"
            sQuery &= "'" & sBPGrpTo & "',"
            sQuery &= "'" & sSlsFr & "',"
            sQuery &= "'" & sSlsTo & "',"
            sQuery &= "'" & sBPCode.Replace("'", "''") & "',"
            sQuery &= "'" & sAsAtDate & "') "
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)

            Return True
            
        Catch ex As Exception
            SBO_Application.MessageBox("[frmARAging].[ExecuteProcedure]:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function ValidateParameter() As Boolean
        Try
            oForm.ActiveItem = "txtBPCode"
            Dim oRecordsetLn As SAPbobsCOM.Recordset = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty
            Dim sQuery As String = String.Empty

            sStart = oForm.DataSources.UserDataSources.Item("txtBPFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtBPTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("BP Code from is greater than BP Code to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtBPFr"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtBPGFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtBPGTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("BP Group from is greater than BP Group to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtBPGFr"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtSlsFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtSlsTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Sales Employee from is greater than Sales Employee to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtSlsFr"
                    Return False
                End If
            End If

            Dim i1 As Integer = 0
            Dim i2 As Integer = 0
            Dim sTemp1 As String = String.Empty
            Dim sTemp2 As String = String.Empty

            iCount = 0
            For iCount = 0 To 4
                sTemp1 = String.Format(sTxtFormat, iCount + 1)
                sTemp2 = String.Format(sValFormat, iCount + 1)
                sTxtBTxt(iCount) = oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx
                sTxtBVal(iCount) = Integer.Parse(oForm.DataSources.UserDataSources.Item(sTemp2).ValueEx)
            Next

            For iCount = 1 To 4
                i1 = sTxtBVal(iCount - 1)
                i2 = sTxtBVal(iCount)
                sTemp1 = String.Format(sValFormat, iCount)
                sTemp2 = String.Format(sValFormat, iCount + 1)

                If (i1 < -1) Then
                    oForm.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[ARAgeing][ValidateParameters] - Value in bucket " & iCount.ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i2 < -1) Then
                    oForm.ActiveItem = sTemp2
                    SBO_Application.StatusBar.SetText("[ARAgeing][ValidateParameters] - Value in bucket " & (iCount + 1).ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i1 > i2) Then
                    oForm.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[ARAgeing][ValidateParameters] - Value in bucket " & (iCount).ToString() & " is greater than value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (iCount = 4) Then
                    If (i1 <> i2) Then
                        oForm.ActiveItem = sTemp1
                        SBO_Application.StatusBar.SetText("[ARAgeing][ValidateParameters] - Value in bucket " & (iCount).ToString() & " is not equal to value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next

            Dim sFileDialog As New SaveFileDialog
            sFileDialog.Filter = "Tab Separated Value|*.xls"
            sFileDialog.Title = "Export To"
            sFileDialog.RestoreDirectory = True
            sFileDialog.CheckFileExists = True
            sFileDialog.DefaultExt = "xls"
            sExcelPath = String.Empty
            bIsExportToExcel = IIf(oForm.DataSources.UserDataSources.Item("chkExcel").ValueEx = "Y", True, False)

            If (bIsExportToExcel) Then
                bIsSaveRunning = True
                bIsCancel = False
                Dim myThread2 As New System.Threading.Thread(AddressOf OpenSaveFileDialog)
                myThread2.SetApartmentState(Threading.ApartmentState.STA)
                myThread2.Start()
                myThread2.Join()
                While (bIsSaveRunning)
                    System.Threading.Thread.CurrentThread.Sleep(500)
                End While
                Return Not bIsCancel
            End If


            SBO_Application.StatusBar.SetText(String.Empty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ARAgeing].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                If pVal.ItemUID = "btnExecute" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                            Return ValidateParameter()
                        End If
                    End If
                End If
            Else
                'AFTER ACTION
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then
                    Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvent = pVal
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oDataTable = oCFLEvent.SelectedObjects
                    If (Not oDataTable Is Nothing) Then
                        Dim sTemp As String = String.Empty
                        Select Case oCFLEvent.ChooseFromListUID
                            Case "cflBPFr"
                                sTemp = oDataTable.GetValue("CardCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtBPFr").ValueEx = sTemp
                                Exit Select
                            Case "cflBPTo"
                                sTemp = oDataTable.GetValue("CardCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtBPTo").ValueEx = sTemp
                                Exit Select
                            Case "CFL_BPCode"
                                sTemp = oDataTable.GetValue("CardCode", 0)
                                oForm.DataSources.UserDataSources.Item("BPCode").ValueEx = sTemp
                                Exit Select
                            Case "CFL_BGFrom"
                                sTemp = oDataTable.GetValue("GroupCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtBPGFr").ValueEx = sTemp
                                Exit Select
                            Case "CFL_BGTo"
                                sTemp = oDataTable.GetValue("GroupCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtBPGTo").ValueEx = sTemp
                                Exit Select
                            Case "CFL_SPFrom"
                                sTemp = oDataTable.GetValue("SlpName", 0)
                                oForm.DataSources.UserDataSources.Item("txtSlsFr").ValueEx = sTemp
                                Exit Select
                            Case "CFL_SPTo"
                                sTemp = oDataTable.GetValue("SlpName", 0)
                                oForm.DataSources.UserDataSources.Item("txtSlsTo").ValueEx = sTemp
                                Exit Select
                            Case "CFL_ProjFr"
                                sTemp = oDataTable.GetValue("PrjCode", 0)
                                oForm.DataSources.UserDataSources.Item("tbProjFr").ValueEx = sTemp
                                Exit Select
                            Case "CFL_ProjTo"
                                sTemp = oDataTable.GetValue("PrjCode", 0)
                                oForm.DataSources.UserDataSources.Item("tbProjTo").ValueEx = sTemp
                                Exit Select
                            Case Else
                                Exit Select
                        End Select
                        Return True
                    End If
                End If

                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    Select Case pVal.ItemUID
                        Case "btnExecute"
                            If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                myThread.SetApartmentState(Threading.ApartmentState.STA)
                                myThread.Start()
                            End If
                        Case "ckFin"
                            PopulateData()
                    End Select
                End If

                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE) Then
                    Select Case pVal.ItemUID
                        Case "txtB4Val"
                            If (pVal.ItemChanged) Then
                                oForm.DataSources.UserDataSources.Item("txtB5Val").ValueEx = oForm.DataSources.UserDataSources.Item("txtB4Val").ValueEx
                            End If
                        Case "txtDate"
                            Dim sVal As String = ""
                            sVal = oForm.DataSources.UserDataSources.Item("ckFin").ValueEx
                            If sVal = "N" Then
                                PopulateData("Y")
                            Else
                                PopulateData("N")
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText("[frmARAgeing].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
