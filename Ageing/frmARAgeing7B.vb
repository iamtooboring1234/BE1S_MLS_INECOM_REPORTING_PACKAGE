Imports System.IO

Public Class frmARAgeing7B

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oItem As SAPbouiCOM.Item
    Private oStatic As SAPbouiCOM.StaticText
    Private oPictureBox As SAPbouiCOM.PictureBox
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Dim oCheck As SAPbouiCOM.CheckBox

    Dim sTxtFormat As String = "txtB{0}txt"
    Dim sValFormat As String = "txtB{0}Val"
    Dim sFTxtFormat As String = "U_Bucket{0}Txt"
    Dim sFValFormat As String = "U_Bucket{0}Val"

    Dim iCount As Integer = 1

    Dim sTxtBTxt As String() = New String(10) {}
    Dim sTxtBVal As Integer() = New Integer(10) {}

    Dim sExcelPath As String = String.Empty
    Dim bIsSaveRunning As Boolean = True
    Dim bIsCancel As Boolean = False

    Dim bIsExportToExcel As Boolean = False
    Dim sRptType As String = String.Empty
    Private g_sARAGERunningDate As String = ""

#End Region

 #Region "Initialization"
    Public Sub ShowForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_ARAGEING_7B.srf") Then

            oForm = SBO_Application.Forms.Item(FRM_ARAGEING_7B)
            oForm.Title = "AR Ageing (7 Buckets) Report"
            oPictureBox = oForm.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"
           
            oForm.Items.Item("lbStyleOpt").TextStyle = 4
            oForm.Items.Item("lbStatus").FontSize = 10

            AddDataSource()
            SetupChooseFromList()
            oForm.Visible = True
        Else
            Try
                oForm = SBO_Application.Forms.Item(FRM_ARAGEING_7B)
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
                .Add("ckReval", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

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
                Dim sTemp As String = String.Empty

                For iCount = 1 To 7
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

                oCheck = oForm.Items.Item("chkExcel").Specific
                oCheck.DataBind.SetBound(True, String.Empty, "chkExcel")
                oCheck.ValOff = "N"
                oCheck.ValOn = "Y"

                oCheck = oForm.Items.Item("ckReval").Specific
                oCheck.DataBind.SetBound(True, String.Empty, "ckReval")
                oCheck.ValOff = "N"
                oCheck.ValOn = "Y"

                oForm.DataSources.UserDataSources.Item("ckReval").ValueEx = "N"
                PopulateData()
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
    Private Sub PopulateData()
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sCurrDate As String = String.Empty
            Dim sQuery As String = String.Empty
            Dim sTemp1 As String = String.Empty

            sTxtBTxt(0) = "0-30"
            sTxtBTxt(1) = "31-60"
            sTxtBTxt(2) = "61-90"
            sTxtBTxt(3) = "91-120"
            sTxtBTxt(4) = "120-150"
            sTxtBTxt(5) = "151-180"
            sTxtBTxt(5) = ">180"

            sTxtBVal(0) = 30
            sTxtBVal(1) = 60
            sTxtBVal(2) = 90
            sTxtBVal(3) = 120
            sTxtBVal(4) = 150
            sTxtBVal(5) = 180
            sTxtBVal(6) = 180

            sQuery = " SELECT * FROM " & oCompany.CompanyDB & ".""@NCM_BUCKET"" WHERE ""U_Type"" = 'NCM_AR_AGEING_7B'"
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                iCount = 0
                For iCount = 0 To 6
                    sTemp1 = String.Format(sFTxtFormat, iCount + 1)
                    sTxtBTxt(iCount) = oRecord.Fields.Item(sTemp1).Value

                    sTemp1 = String.Format(sFValFormat, iCount + 1)
                    sTxtBVal(iCount) = oRecord.Fields.Item(sTemp1).Value
                Next
            End If

            For iCount = 0 To 6
                sTemp1 = String.Format(sTxtFormat, iCount + 1)
                oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBTxt(iCount)

                sTemp1 = String.Format(sValFormat, iCount + 1)
                oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBVal(iCount)
            Next

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
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition

        Try
            oCFLs = oForm.ChooseFromLists
            'Production Order Choose from list
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "2"
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
        Catch ex As Exception
            Throw New Exception("[ARAgeing].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Function IsSharedFileExist() As Boolean
        Try
            '' 1 - File not found thus use local - 0 - File's found thus use shared file - [-1] - error
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim RptType As String = String.Empty

            oCombo = oForm.Items.Item("cboRptType").Specific
            RptType = oCombo.Selected.Value
            g_sReportFilename = ""

            Select Case RptType
                Case 0
                    g_sReportFilename = GetSharedFilePath(ReportName.ARAging7B_Details)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
                Case 1
                    g_sReportFilename = GetSharedFilePath(ReportName.ARAging7B_Summary)
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
    Private Function GetLocalCurrency() As String
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""MainCurncy"" FROM " & oCompany.CompanyDB & ".""OADM"" ")
            If (oRec.RecordCount > 0) Then
                Return oRec.Fields.Item(0).Value.ToString()
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GetLocalCurrency] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return String.Empty
    End Function
    Private Function ExecuteProcedure() As Boolean
        Dim sBPCode As String = ""
        Dim sAsAtDate As String = ""
        Dim dtAsAtDate As DateTime
        Dim sBPCodeFr As String = String.Empty
        Dim sBPCodeTo As String = String.Empty
        Dim sBPGrpFr As String = String.Empty
        Dim sBPGrpTo As String = String.Empty
        Dim sSlsFr As String = String.Empty
        Dim sSlsTo As String = String.Empty

        Try
            g_bIsShared = IsSharedFileExist()

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
            'BPCode = IIf(oEdit.Value.Trim = "", "%", "%" & oEdit.Value.Trim.Replace("*", "%") & "%")
            sBPCode = oEdit.Value
            oEdit = oForm.Items.Item("txtDate").Specific
            sAsAtDate = oEdit.Value
            'dtAsAtDate = New DateTime(Left(AsAtDate, 4), Mid(AsAtDate, 5, 2), Right(AsAtDate, 2))

            oStatic = oForm.Items.Item("lbStatus").Specific
            oStatic.Caption = "Executing Store Procedure. Please wait..."

            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sQuery = " CALL SP_AR_AGEING ("
            sQuery &= "'" & g_sARAGERunningDate & oCompany.UserName & "',"
            sQuery &= "'" & sBPCodeFr.Replace("'", "''") & "',"
            sQuery &= "'" & sBPCodeTo.Replace("'", "''") & "',"
            sQuery &= "'" & sBPGrpFr & "',"
            sQuery &= "'" & sBPGrpTo & "',"
            sQuery &= "'" & sSlsFr & "',"
            sQuery &= "'" & sSlsTo & "',"
            sQuery &= "'" & sBPCode.Replace("'", "''") & "',"
            sQuery &= "'" & sAsAtDate & "')"
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

            Dim sTemp1 As String = String.Empty
            Dim sTemp2 As String = String.Empty
            iCount = 0
            For iCount = 0 To 6
                sTemp1 = String.Format(sTxtFormat, iCount + 1)
                sTemp2 = String.Format(sValFormat, iCount + 1)
                sTxtBTxt(iCount) = oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx
                sTxtBVal(iCount) = Integer.Parse(oForm.DataSources.UserDataSources.Item(sTemp2).ValueEx)
            Next

            Dim i1 As Integer = 0
            Dim i2 As Integer = 0

            For iCount = 1 To 6
                i1 = sTxtBVal(iCount - 1)
                i2 = sTxtBVal(iCount)
                sTemp1 = String.Format(sValFormat, iCount)
                sTemp2 = String.Format(sValFormat, iCount + 1)

                If (i1 < -1) Then
                    oForm.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[ValidateParameters] - Value in bucket " & iCount.ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i2 < -1) Then
                    oForm.ActiveItem = sTemp2
                    SBO_Application.StatusBar.SetText("[ValidateParameters] - Value in bucket " & (iCount + 1).ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i1 > i2) Then
                    oForm.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[ValidateParameters] - Value in bucket " & (iCount).ToString() & " is greater than value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (iCount = 6) Then
                    If (i1 <> i2) Then
                        oForm.ActiveItem = sTemp1
                        SBO_Application.StatusBar.SetText("[ValidateParameters] - Value in bucket " & (iCount).ToString() & " is not equal to value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        oForm.Items.Item("btnExecute").Enabled = False
        Try
            Dim frm As New Hydac_FormViewer
            Dim bIsContinue As Boolean = False
            Dim oTest As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)) FROM DUMMY")
            If oTest.RecordCount > 0 Then
                oTest.MoveFirst()
                g_sARAGERunningDate = Convert.ToString(oTest.Fields.Item(0).Value)
            End If
            oTest = Nothing

            Try
                If (ExecuteProcedure()) Then
                    bIsContinue = True
                    oEdit = oForm.Items.Item("txtDate").Specific
                    Dim AsAtDate As String = oEdit.Value
                    oCombo = oForm.Items.Item("cboAgeBy").Specific
                    Dim AgeingBy As String = oCombo.Selected.Value
                    Dim sAgeingBy As String = oCombo.Selected.Description

                    oCombo = oForm.Items.Item("cboRptType").Specific
                    Dim RptType As String = oCombo.Selected.Value

                    Dim sSplitReval As String = "N"
                    Dim sBPCodeFr As String = String.Empty
                    Dim sBPCodeTo As String = String.Empty
                    Dim sBPGrpFr As String = String.Empty
                    Dim sBPGrpTo As String = String.Empty
                    Dim sSlsFr As String = String.Empty
                    Dim sSlsTo As String = String.Empty
                    Dim sLocalCurr As String = String.Empty
                    Dim iPageBreak As Integer = 0

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

                    oCheck = DirectCast(oForm.Items.Item("ckReval").Specific, SAPbouiCOM.CheckBox)
                    If (oCheck.Checked) Then
                        sSplitReval = "Y"
                    Else
                        sSplitReval = "N"
                    End If

                    SBO_Application.StatusBar.SetText("Opening Report. Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Select Case RptType
                        Case 0
                            frm.Text = "AR Ageing 7 Buckets Report [Detail]"
                            frm.ReportType = AgeingType.ARAgeing
                        Case 1
                            frm.Text = "AR Ageing 7 Buckets Report [Summary]"
                            frm.ReportType = AgeingType.ARAgeingSummary
                            If sSplitReval = "N" Then
                                ' update [@ncm_arageing]
                                SBO_Application.StatusBar.SetText("Opening Report. Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                Dim sLoop As String = ""
                                Dim sSelect As String = ""
                                Dim sUpdate As String = ""
                                Dim sCurrency As String = ""
                                Dim oSelect As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oUpdate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim oLoop As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                sSelect = " select A.""CARDCODE"", COUNT(A.""DOCCUR"") AS ""CountTotal"" FROM "
                                sSelect &= " (select T1.""CARDCODE"", T1.""DOCCUR"" "
                                sSelect &= "  from " & oCompany.CompanyDB & ".""@NCM_AR_AGEING"" T1 "
                                sSelect &= "  where T1.""USERNAME"" = '" & g_sARAGERunningDate & oCompany.UserName.Trim & "'"
                                sSelect &= "  group by T1.""CARDCODE"", T1.""DOCCUR"" ) A"
                                sSelect &= " GROUP BY A.""CARDCODE"" HAVING COUNT(A.""DOCCUR"") > 1"
                                sSelect &= " ORDER BY A.""CARDCODE"" "
                                oSelect.DoQuery(sSelect)
                                If oSelect.RecordCount > 0 Then
                                    oSelect.MoveFirst()
                                    While Not oSelect.EoF
                                        sLoop = "  SELECT CASE WHEN ""Currency"" = '##' THEN '" & sLocalCurr & "' ELSE ""Currency"" END "
                                        sLoop &= " FROM " & oCompany.CompanyDB & ".""OCRD"" "
                                        sLoop &= " WHERE ""CardCode"" = '" & oSelect.Fields.Item(0).Value & "'"
                                        oLoop = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oLoop.DoQuery(sLoop)
                                        If oLoop.RecordCount > 0 Then
                                            oLoop.MoveFirst()
                                            sCurrency = oLoop.Fields.Item(0).Value

                                            sUpdate = " UPDATE " & oCompany.CompanyDB & ".""@NCM_AR_AGEING"" "
                                            sUpdate &= " SET ""DOCCUR"" = '" & sCurrency & "' "
                                            sUpdate &= " WHERE ""USERNAME"" = '" & g_sARAGERunningDate & oCompany.UserName.Trim & "'"
                                            sUpdate &= " AND ""CARDCODE"" = '" & oSelect.Fields.Item(0).Value & "'"
                                            oUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oUpdate.DoQuery(sUpdate)

                                        End If
                                        oSelect.MoveNext()
                                    End While
                                End If
                                oSelect = Nothing
                                oUpdate = Nothing
                                oLoop = Nothing
                            End If
                    End Select

                    frm.IsShared = g_bIsShared
                    frm.SharedReportName = g_sReportFilename
                    frm.ReportName = ReportName.ARAging7B_Details
                    frm.ARAGERunningDate = g_sARAGERunningDate & oCompany.UserName
                    frm.DBPasswordViewer = DBPassword
                    frm.DBUsernameViewer = DBUsername
                    frm.Username = oCompany.UserName
                    frm.AsAtDate = AsAtDate
                    frm.AgeBy = AgeingBy
                    frm.BPCode = sBPCode
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
            SBO_Application.MessageBox("[ARAgeing].[LoadViewer]:" & ex.ToString)
        Finally
        End Try
    End Sub
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
                            Case Else
                                Exit Select
                        End Select
                        Return True
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.ItemUID = "btnExecute" Then
                        If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                            myThread.Start()
                        End If
                    End If
                End If
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE) Then
                    Select Case pVal.ItemUID
                        Case "txtB6Val"
                            If (pVal.ItemChanged) Then
                                oForm.DataSources.UserDataSources.Item("txtB7Val").ValueEx = oForm.DataSources.UserDataSources.Item("txtB6Val").ValueEx
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ARAgeing].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class