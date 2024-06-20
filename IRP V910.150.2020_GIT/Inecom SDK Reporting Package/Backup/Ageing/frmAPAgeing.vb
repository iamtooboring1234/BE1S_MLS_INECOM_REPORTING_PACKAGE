Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class frmAPAgeing

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oItem As SAPbouiCOM.Item
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox

    Private oStatic As SAPbouiCOM.StaticText
    Private oPictureBox As SAPbouiCOM.PictureBox
    Private oMtrxAPA As SAPbouiCOM.Matrix

    Private g_sReportFilename As String = ""
    Private g_StructureFilename As String = ""
    Private g_bIsShared As Boolean = False
    Private ds As DataSet

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
    Private g_sAll As String = "N"
    Private g_sAPAGERunningDate As String = ""
#End Region

#Region "Initialization"
    Public Sub ShowForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmAPAgeing.srf") Then
            oForm = SBO_Application.Forms.Item("ncmAPAgeing")
            oPictureBox = oForm.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"

            ' COMMENTED OUT - BUSINESS UNIT CODE
            oForm.Items.Item("lbDim2").Visible = False
            oForm.Items.Item("mxAPA").Visible = False
            oForm.Items.Item("btSelect").Visible = False
            oForm.Items.Item("btDeselect").Visible = False


            AddDataSource()
            SetupChooseFromList()
            g_sAll = "N"
            oForm.Visible = True
        Else
            Try
                oForm = SBO_Application.Forms.Item("ncmAPAgeing")
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
            Dim oCols As SAPbouiCOM.Columns
            Dim oColn As SAPbouiCOM.Column
            Dim sTemp As String = String.Empty

            oMtrxAPA = oForm.Items.Item("mxAPA").Specific
            With oForm.DataSources.UserDataSources
                .Add("xRow", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
                .Add("xSel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("xBus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
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
            End With

            oCols = oMtrxAPA.Columns
            oColn = oCols.Item("cRow")
            oColn.DataBind.SetBound(True, , "xRow")
            oColn = oCols.Item("cBus")
            oColn.DataBind.SetBound(True, , "xBus")
            oColn = oCols.Item("cSel")
            oColn.DataBind.SetBound(True, , "xSel")
            oColn.ValOff = "N"
            oColn.ValOn = "Y"

            oEdit = oForm.Items.Item("txtBPCode").Specific
            oEdit.DataBind.SetBound(True, "", "BPCode")
            oEdit = oForm.Items.Item("txtDate").Specific
            oEdit.DataBind.SetBound(True, "", "dDate")
            oForm.DataSources.UserDataSources.Item("dDate").ValueEx = Now.ToString("yyyyMMdd")

            oCombo = oForm.Items.Item("cboAgeBy").Specific
            oCombo.ValidValues.Add("0", "Document Date")
            oCombo.ValidValues.Add("1", "Due Date")
            oCombo.ValidValues.Add("2", "Posting Date")
            oCombo.DataBind.SetBound(True, "", "AgeBy")
            oForm.DataSources.UserDataSources.Item("AgeBy").ValueEx = "0"

            oCombo = oForm.Items.Item("cboRptType").Specific
            oCombo.ValidValues.Add("0", "Details")
            oCombo.ValidValues.Add("1", "Summary")
            oCombo.DataBind.SetBound(True, "", "RptType")
            oForm.DataSources.UserDataSources.Item("RptType").ValueEx = "0"

            For iCount = 1 To 5
                oForm.DataSources.UserDataSources.Add(String.Format(sTxtFormat, iCount), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                oForm.DataSources.UserDataSources.Add(String.Format(sValFormat, iCount), SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
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

            oCheck = oForm.Items.Item("chkExcel").Specific
            oCheck.DataBind.SetBound(True, String.Empty, "chkExcel")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"

            oCheck = DirectCast(oForm.Items.Item("chkPage").Specific, SAPbouiCOM.CheckBox)
            oCheck.DataBind.SetBound(True, String.Empty, "chkPage")
            oCheck.ValOff = "0"
            oCheck.ValOn = "1"
            PopulateData()

        Catch ex As Exception
            SBO_Application.MessageBox("[frmAPAgeing].[AddDataSource]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub PopulateData()
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sCurrDate As String = String.Empty
            Dim sQuery As String = String.Empty
            Dim sTemp1 As String = String.Empty
            Dim iRow As Integer = 1

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

            ' ------------------------------------------------------------------
            ' BUSINESS UNIT CODE
            'sQuery = "  SELECT U_Dim2 FROM [OCRD] "
            'sQuery &= " WHERE ISNULL(U_Dim2,'') <> ''  AND CardType = 'S' "
            'sQuery &= " GROUP BY U_Dim2 "
            'oRec.DoQuery(sQuery)
            'If oRec.RecordCount > 0 Then
            '    oRec.MoveFirst()
            '    While Not oRec.EoF
            '        With oForm.DataSources.UserDataSources
            '            .Item("xRow").ValueEx = iRow
            '            .Item("xSel").ValueEx = "N"
            '            .Item("xBus").ValueEx = oRec.Fields.Item(0).Value
            '        End With

            '        iRow += 1
            '        oMtrxAPA.AddRow()
            '        oRec.MoveNext()
            '    End While
            'End If
            'oRec = Nothing
            ' ------------------------------------------------------------------

            sQuery = " SELECT * FROM ""@NCM_BUCKET"" WHERE ""U_Type"" = 'NCM_AP_AGEING'"
            oRecord.DoQuery(sQuery)
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
    Private Sub ShowStatus(ByVal sStatus As String)
        Dim oStaticText As SAPbouiCOM.StaticText = oForm.Items.Item("lbStatus").Specific
        oStaticText.Caption = sStatus
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
            oCon.CondVal = "S"
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
            oCon.CondVal = "S"
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
            oCon.CondVal = "S"
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
            oCon.CondVal = "S"
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
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oForm.Items.Item("txtBPGFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BGFrom"
            oEditLn.ChooseFromListAlias = "GroupCode"
            ' ----------------------------------------
        Catch ex As Exception
            Throw New Exception("[APAgeing].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Function IsSharedFileExist() As Boolean
        Try
            '' 1 - File not found thus use local - 0 - File's found thus use shared file - [-1] - error
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""
            Dim RptType As String = String.Empty
            oCombo = oForm.Items.Item("cboRptType").Specific
            RptType = oCombo.Selected.Value
            g_sReportFilename = ""
            g_StructureFilename = ""

            Select Case RptType
                Case 0
                    sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
                    sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.APAging_Details) & "'"

                    g_sReportFilename = GetSharedFilePath(ReportName.APAging_Details)
                   
                Case 1
                    sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
                    sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.APAging_Summary) & "'"

                    g_sReportFilename = GetSharedFilePath(ReportName.APAging_Summary)
            End Select

            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    'okay
                End If
            End If

            Dim sCheck As String = ""
            Dim oCheck As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sCheck = "  SELECT ""OBJECT_NAME"" FROM SYS.OBJECTS  "
            sCheck &= " WHERE ""SCHEMA_NAME"" = '" & oCompany.CompanyDB & "' "
            sCheck &= " AND ""OBJECT_TYPE"" = 'TABLE' "
            sCheck &= " AND ""OBJECT_NAME"" ='@NCM_RPT_STRUCTURE' "
            oCheck.DoQuery(sCheck)
            If oCheck.RecordCount > 0 Then
                oCheck = Nothing

                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    g_StructureFilename = oRec.Fields.Item(0).Value.ToString
                    If File.Exists(g_StructureFilename) = False Then
                        g_StructureFilename = ""
                    End If
                End If
            Else
                oCheck = Nothing
            End If

            Return True
        Catch ex As Exception
            g_sReportFilename = ""
            g_StructureFilename = ""
            SBO_Application.StatusBar.SetText("[APAgeing].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function PrepareDataset() As Boolean
        Try
            If g_StructureFilename.Length <= 0 Then
                ds = New DS_AGEING

            Else
                ds = New DataSet
                ds.ReadXml(g_StructureFilename)
            End If

            Dim ProviderName As String = "System.Data.Odbc"
            Dim sQuery As String = ""
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            Dim dtAGE As System.Data.DataTable
            Dim dtOCRD As System.Data.DataTable

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            '--------------------------------------------------------
            'OADM (Company Details)
            '--------------------------------------------------------
            sQuery = "  SELECT T1.""CardCode"", T1.""CardName"", IFNULL(T1.""GroupCode"",0) ""GroupCode"", IFNULL(T2.""GroupName"",'') ""GroupName"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OCRD"" T1 "
            sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T2 "
            sQuery &= " ON T1.""GroupCode"" = T2.""GroupCode"" "
            sQuery &= " WHERE T1.""CardType"" = 'S' "

            dtOCRD = ds.Tables("OCRD")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOCRD)

            '--------------------------------------------------------
            'NCM_AR_AGEING
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""@NCM_AP_AGEING"" "
            sQuery &= " WHERE ""USERNAME"" = '" & g_sAPAGERunningDate & oCompany.UserName & "' "
            dtAGE = ds.Tables("@NCM_AP_AGEING")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtAGE)

            dbConn.Close()

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PrepareDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub LoadViewer()
        oForm.Items.Item("btnExecute").Enabled = False
        Dim frm As New Hydac_FormViewer
        Dim bIsContinue As Boolean = False
        Dim sTempDirectory As String = ""
        Dim sPathFormat As String = "{0}\APAGEING_{1}.pdf"
        Dim sCurrDate As String = ""
        Dim sFinalExportPath As String = ""
        Dim sFinalFileName As String = ""
        Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")

        Try
            Try
                Dim oTest As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTest.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD')  FROM DUMMY")
                If oTest.RecordCount > 0 Then
                    oTest.MoveFirst()
                    g_sAPAGERunningDate = Convert.ToString(oTest.Fields.Item(0).Value)
                    sCurrDate = Convert.ToString(oTest.Fields.Item(1).Value).Trim
                End If
                oTest = Nothing

                ' g_sAPAGERunningDate = "2016-01-14 16:02:55.8120"
                ' ===============================================================================
                ' get the folder of AR SOA of the current DB Name
                ' set to local
                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\APAGEING\" & oCompany.CompanyDB
                Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                If Not di.Exists Then
                    di.Create()
                End If
                sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                sFinalFileName = di.FullName & "\APAGEING_" & sCurrDate & "_" & sCurrTime & ".pdf"
                ' ===============================================================================

                If ExecuteProcedure() Then
                    If PrepareDataset() Then
                        bIsContinue = True
                        oEdit = oForm.Items.Item("txtDate").Specific
                        Dim AsAtDate As String = oEdit.Value
                        oCombo = oForm.Items.Item("cboAgeBy").Specific
                        Dim AgeingBy As String = oCombo.Selected.Value
                        Dim sAgeingBy As String = oCombo.Selected.Description

                        oCombo = oForm.Items.Item("cboRptType").Specific
                        Dim RptType As String = oCombo.Selected.Value

                        Dim sBPCodeFr As String = ""
                        Dim sBPCodeTo As String = ""
                        Dim sBPGrpFr As String = ""
                        Dim sBPGrpTo As String = ""
                        Dim sSlsFr As String = ""
                        Dim sSlsTo As String = ""
                        Dim sLocalCurr As String = ""
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

                        SBO_Application.StatusBar.SetText("Opening Report. Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Select Case RptType
                            Case 0
                                frm.Text = "AP Ageing Details Report"
                                frm.ReportType = AgeingType.APAgeing
                            Case 1
                                frm.Text = "AP Ageing Summary Report"
                                frm.ReportType = AgeingType.APAgeingSummary
                        End Select

                        frm.Dataset = ds
                        frm.ReportName = ReportName.APAging_Details
                        frm.DBPasswordViewer = DBPassword
                        frm.DBUsernameViewer = DBUsername
                        frm.IsShared = g_bIsShared
                        frm.SharedReportName = g_sReportFilename
                        frm.APAGERunningDate = g_sAPAGERunningDate & oCompany.UserName
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

                        frm.ExportPath = sFinalFileName
                        Select Case SBO_Application.ClientType
                            Case SAPbouiCOM.BoClientType.ct_Desktop
                                frm.ClientType = "D"
                            Case SAPbouiCOM.BoClientType.ct_Browser
                                frm.ClientType = "S"
                        End Select

                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item("btnExecute").Enabled = True
            End Try
            If bIsContinue Then
                Select Case SBO_Application.ClientType
                    Case SAPbouiCOM.BoClientType.ct_Desktop
                        If (bIsExportToExcel) Then
                            'Export To Excel
                            '--------------------------------------------------------------------------------
                            frm.OPEN_HANADS_AGEING_5BUCKETS()
                        Else
                            'Not Export To Excel
                            '--------------------------------------------------------------------------------
                            frm.ShowDialog()
                        End If

                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.OPEN_HANADS_AGEING_5BUCKETS()
                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
    Private Function ExecuteProcedure() As Boolean
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sQuery As String = ""
        Dim sBPCode As String = ""
        Dim sAsAtDate As String = ""
        Dim dtAsAtDate As DateTime
        Dim sBPCodeFr As String = String.Empty
        Dim sBPCodeTo As String = String.Empty
        Dim sBPGrpFr As String = String.Empty
        Dim sBPGrpTo As String = String.Empty
        Dim sSlsFr As String = String.Empty
        Dim sSlsTo As String = String.Empty
        'Dim sSelected As String = GetSelected()

        g_bIsShared = IsSharedFileExist()
        If (g_bIsShared) Then
            If g_sReportFilename.Trim.Length > 0 Then
                If (Not File.Exists(g_sReportFilename)) Then
                    g_bIsShared = False
                    g_sReportFilename = ""
                End If
            Else
                g_bIsShared = False
                g_sReportFilename = ""
            End If
        End If

        ' Return True

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
            sBPCode = oEdit.Value.Trim
            oEdit = oForm.Items.Item("txtDate").Specific
            sAsAtDate = oEdit.Value
            dtAsAtDate = New DateTime(Left(sAsAtDate, 4), Mid(sAsAtDate, 5, 2), Right(sAsAtDate, 2))

            oStatic = oForm.Items.Item("lbStatus").Specific
            oStatic.Caption = "Executing Store Procedure. Please wait..."

            sQuery = "CALL SP_AP_AGEING ("
            sQuery &= "'" & g_sAPAGERunningDate & oCompany.UserName & "',"
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
            SBO_Application.MessageBox("[APAgeing][ExecuteProcedure]:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function ValidateParameter() As Boolean
        Try
            oForm.ActiveItem = "txtBPCode"
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty
            Dim sQuery As String = String.Empty

            With oForm.DataSources.UserDataSources
                sStart = .Item("txtBPFr").ValueEx.Replace("'", "''")
                sEnd = .Item("txtBPTo").ValueEx.Replace("'", "''")
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("BP Code From is greater than BP Code To.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.ActiveItem = "txtBPFr"
                        Return False
                    End If
                End If

                sStart = .Item("txtBPGFr").ValueEx.Replace("'", "''")
                sEnd = .Item("txtBPGTo").ValueEx.Replace("'", "''")
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("BP Group From is greater than BP Group To.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.ActiveItem = "txtBPGFr"
                        Return False
                    End If
                End If

                sStart = .Item("txtSlsFr").ValueEx.Replace("'", "''")
                sEnd = .Item("txtSlsTo").ValueEx.Replace("'", "''")
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("Sales Employee From is greater than Sales Employee To.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.ActiveItem = "txtSlsFr"
                        Return False
                    End If
                End If
            End With

            Dim sTemp1 As String = String.Empty
            Dim sTemp2 As String = String.Empty
            iCount = 0
            For iCount = 0 To 4
                sTemp1 = String.Format(sTxtFormat, iCount + 1)
                sTemp2 = String.Format(sValFormat, iCount + 1)
                sTxtBTxt(iCount) = oForm.DataSources.UserDataSources.Item(sTemp1).ValueEx
                sTxtBVal(iCount) = Integer.Parse(oForm.DataSources.UserDataSources.Item(sTemp2).ValueEx)
            Next

            Dim i1 As Integer = 0
            Dim i2 As Integer = 0

            For iCount = 1 To 4
                i1 = sTxtBVal(iCount - 1)
                i2 = sTxtBVal(iCount)
                sTemp1 = String.Format(sValFormat, iCount)
                sTemp2 = String.Format(sValFormat, iCount + 1)
                If (i1 < -1) Then
                    oForm.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[AP_Ageing][ValidateParameters] - Value in bucket " & iCount.ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i2 < -1) Then
                    oForm.ActiveItem = sTemp2
                    SBO_Application.StatusBar.SetText("[AP_Ageing][ValidateParameters] - Value in bucket " & (iCount + 1).ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i1 > i2) Then
                    oForm.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[AP_Ageing][ValidateParameters] - Value in bucket " & (iCount).ToString() & " is greater than value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (iCount = 4) Then
                    If (i1 <> i2) Then
                        oForm.ActiveItem = sTemp1
                        SBO_Application.StatusBar.SetText("[AP_Ageing][ValidateParameters] - Value in bucket " & (iCount).ToString() & " is not equal to value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

            Return True

        Catch ex As Exception
            SBO_Application.MessageBox("[APAgeing].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
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
            oRec = Nothing
            Return ""
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[APAgeing].[GetLocalCurrency] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return String.Empty
    End Function
    Private Sub SelectAll(ByVal sCol As String, ByVal sVal As String)
        Try
            oForm.Freeze(True)
            For i As Integer = 1 To oMtrxAPA.VisualRowCount Step 1
                oMtrxAPA.GetLineData(i)
                oForm.DataSources.UserDataSources.Item(sCol).ValueEx = sVal
                oMtrxAPA.SetLineData(i)
            Next
            oForm.Freeze(False)

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[APAgeing].[SelectAll] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
    Private Function GetSelected() As String
        Try
            Dim sSelect As String = ""
            For i As Integer = 1 To oMtrxAPA.VisualRowCount Step 1
                oMtrxAPA.GetLineData(i)
                If oForm.DataSources.UserDataSources.Item("xSel").ValueEx = "Y" Then
                    sSelect &= "'" & oForm.DataSources.UserDataSources.Item("xBus").ValueEx & "',"
                End If
            Next
            If sSelect.Length > 0 Then
                sSelect = sSelect.Remove(sSelect.Length - 1, 1)
            End If
            Return sSelect
        Catch ex As Exception
            Return ""
            SBO_Application.StatusBar.SetText("[APAgeing].GetSelected] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "btnExecute" Then
                            If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                Return ValidateParameter()
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Select Case pVal.ItemUID
                            Case "btSelect"
                                If oForm.Items.Item("btnExecute").Enabled = True Then
                                    g_sAll = "Y"
                                    SelectAll("xSel", g_sAll)
                                End If
                            Case "btDeselect"
                                If oForm.Items.Item("btnExecute").Enabled = True Then
                                    g_sAll = "N"
                                    SelectAll("xSel", g_sAll)
                                End If
                        End Select
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
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
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "btnExecute" Then
                            If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                myThread.SetApartmentState(Threading.ApartmentState.STA)
                                myThread.Start()
                            End If
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        Select Case pVal.ItemUID
                            Case "txtB4Val"
                                If (pVal.ItemChanged) Then
                                    oForm.DataSources.UserDataSources.Item("txtB5Val").ValueEx = oForm.DataSources.UserDataSources.Item("txtB4Val").ValueEx
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText("[APAgeing].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class