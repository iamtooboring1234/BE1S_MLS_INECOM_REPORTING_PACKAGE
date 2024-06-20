Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class frmApSOA

#Region "Global Variables"
    Private sQuery As String = ""
    Private BPCode As String
    Private AsAtDate As DateTime
    Private FromDate As DateTime
    Private IsBBF As String
    Private IsGAT As String
    Private dsSOA As DataSet

    Private g_StructureFilename As String = ""
    Private g_sReportFilename As String = ""
    Private g_bIsShared As Boolean = False
    Private g_sAPSOARunningDate As String = ""

    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
#End Region

#Region "Intialize Application"
    Public Sub New()
        Try
            'Setup_Notes()
        Catch ex As Exception
            MsgBox("[frmApSoa].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        Dim oPictureBox As SAPbouiCOM.PictureBox
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmSOA_AP.srf") Then

            oForm = SBO_Application.Forms.Item("ncmSOA_AP")
            oPictureBox = oForm.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"

            oForm.Items.Item("lbStatus").FontSize = 10
            oForm.Items.Item("lbStyleOpt").TextStyle = 4

            AddDataSource()
            SetDatasource()
            Retrieve_Notes()
            SetupChooseFromList()

            oForm.Items.Item("etBPCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Visible = True
        Else
            Try
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
        With oForm.DataSources.UserDataSources
            .Add("DateType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            .Add("DateAsAt", SAPbouiCOM.BoDataType.dt_DATE)
            .Add("txtDateFr", SAPbouiCOM.BoDataType.dt_DATE)
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
            .Add("txtBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            .Add("txtBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            .Add("txtBPGFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtBPGTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtSlsFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtSlsTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        End With
    End Sub
    Private Sub SetDatasource()
        oEdit = oForm.Items.Item("etBPCode").Specific
        oEdit.DataBind.SetBound(True, "", "BPCode")
        oEdit = oForm.Items.Item("etDateAsAt").Specific
        oEdit.DataBind.SetBound(True, "", "DateAsAt")
        oEdit.Value = Now.ToString("yyyyMMdd")
        oEdit = oForm.Items.Item("txtDateFr").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtDateFr")
        oEdit.Value = Now.ToString("yyyyMM01")
        oEdit = oForm.Items.Item("etNotes").Specific
        oEdit.DataBind.SetBound(True, "", "Notes")

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

        oCombo = oForm.Items.Item("cbDateType").Specific
        oCombo.ValidValues.Add("0", "Document Date")
        oCombo.ValidValues.Add("1", "Due Date")
        oCombo.ValidValues.Add("2", "Posting Date")
        oCombo.DataBind.SetBound(True, "", "DateType")
        oForm.DataSources.UserDataSources.Item("DateType").ValueEx = "0"

        oCombo = oForm.Items.Item("cbPrdType").Specific
        oCombo.ValidValues.Add("0", "Every 30 Days")
        oCombo.ValidValues.Add("1", "Every Month")
        oCombo.DataBind.SetBound(True, "", "Period")
        oForm.DataSources.UserDataSources.Item("Period").ValueEx = "0"

        oCheck = oForm.Items.Item("ckLogo").Specific
        oCheck.DataBind.SetBound(True, "", "Logo")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
        oCheck = oForm.Items.Item("ckHDR").Specific
        oCheck.DataBind.SetBound(True, "", "HDR")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
        oCheck = oForm.Items.Item("ckBBF").Specific
        oCheck.DataBind.SetBound(True, "", "BBF")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
        oCheck = oForm.Items.Item("ckSNP").Specific
        oCheck.DataBind.SetBound(True, "", "SNP")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
        oCheck = oForm.Items.Item("ckGAT").Specific
        oCheck.DataBind.SetBound(True, "", "GAT")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
        oCheck = oForm.Items.Item("ckHAS").Specific
        oCheck.DataBind.SetBound(True, "", "HAS")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
        oCheck = oForm.Items.Item("ckHFN").Specific
        oCheck.DataBind.SetBound(True, "", "HFN")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"

        oCheck = oForm.Items.Item("ckExc").Specific
        oCheck.DataBind.SetBound(True, "", "EXC")
        oCheck.ValOff = "N"
        oCheck.ValOn = "Y"
    End Sub
    Private Sub ShowStatus(ByVal sStatus As String)
        Dim oStaticText As SAPbouiCOM.StaticText = oForm.Items.Item("lbStatus").Specific
        oStaticText.Caption = sStatus
    End Sub
    Private Function IsSharedFileExist() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            g_sReportFilename = ""
            g_StructureFilename = ""

            sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.APSoa) & "'"

            g_sReportFilename = GetSharedFilePath(ReportName.APSoa)
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
            SBO_Application.StatusBar.SetText("[APSOA].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        oForm.Items.Item("btnExecute").Enabled = False
        Try
            Dim frm As New Hydac_FormViewer
            Dim bIsContinue As Boolean = False
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\APSOA_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")

            Try
                g_sAPSOARunningDate = ""
                Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    g_sAPSOARunningDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
                    sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
                End If
                oRec = Nothing

                'g_sAPSOARunningDate = "2015-02-24 16:52:20.4710"
                ' ===============================================================================
                ' get the folder of AR SOA of the current DB Name
                ' set to local
                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\APSOA\" & oCompany.CompanyDB
                Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                If Not di.Exists Then
                    di.Create()
                End If
                sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                sFinalFileName = di.FullName & "\APSOA_" & sCurrDate & "_" & sCurrTime & ".pdf"
                ' ===============================================================================

                If Save_Settings() Then
                    If ExecuteProcedure() Then
                        If PrepareDataset() Then
                            ' ==========================================================
                            oCombo = oForm.Items.Item("cbDateType").Specific
                            If oCombo.Selected Is Nothing Then
                                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                            frm.Report = oCombo.Selected.Value

                            oCombo = oForm.Items.Item("cbPrdType").Specific
                            If oCombo.Selected Is Nothing Then
                                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                            frm.Period = oCombo.Selected.Value

                            frm.Dataset = dsSOA
                            frm.DatabaseServer = oCompany.Server
                            frm.DatabaseName = oCompany.CompanyDB
                            frm.APSOARunningDate = g_sAPSOARunningDate & oCompany.UserName
                            frm.IsShared = g_bIsShared
                            frm.SharedReportName = g_sReportFilename
                            frm.DBUsernameViewer = DBUsername
                            frm.DBPasswordViewer = DBPassword
                            frm.Username = oCompany.UserName
                            frm.AsAtDate = AsAtDate.ToString("yyyyMMdd")
                            frm.ReportName = ReportName.APSoa
                            frm.ExportPath = sFinalFileName
                            Select Case SBO_Application.ClientType
                                Case SAPbouiCOM.BoClientType.ct_Desktop
                                    frm.ClientType = "D"
                                Case SAPbouiCOM.BoClientType.ct_Browser
                                    frm.ClientType = "S"
                            End Select

                            oCheck = oForm.Items.Item("ckLogo").Specific
                            frm.HideLogo = IIf(oCheck.Checked, True, False)
                            oCheck = oForm.Items.Item("ckHDR").Specific
                            frm.HideHeader = IIf(oCheck.Checked, True, False)
                            oCheck = oForm.Items.Item("ckBBF").Specific
                            frm.IsBBF = IIf(oCheck.Checked, 1, 0)
                            oCheck = oForm.Items.Item("ckSNP").Specific
                            frm.IsSNP = IIf(oCheck.Checked, 1, 0)
                            oCheck = oForm.Items.Item("ckGAT").Specific
                            frm.IsGAT = IIf(oCheck.Checked, 1, 0)
                            oCheck = oForm.Items.Item("ckHAS").Specific
                            frm.IsHAS = IIf(oCheck.Checked, 1, 0)
                            oCheck = oForm.Items.Item("ckHFN").Specific
                            frm.IsHFN = IIf(oCheck.Checked, 1, 0)

                            bIsContinue = True
                            ' ==========================================================
                        End If
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
                        frm.ShowDialog()

                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.OPEN_HANADS_APSOA()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[APSOA].[LoadViewer]:" & ex.Message)
        Finally
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

            oEditLn = DirectCast(oForm.Items.Item("etBPCode").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BPCode"
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
            oCon.CondVal = "S"
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
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            oEditLn = DirectCast(oForm.Items.Item("txtBPGTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BGTo"
            oEditLn.ChooseFromListAlias = "GroupCode"
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
        Catch ex As Exception
            Throw New Exception("[APSOA].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function PrepareDataset() As Boolean
        Try
            If g_StructureFilename.Length <= 0 Then
                dsSOA = New DS_SOA
            Else
                dsSOA = New DataSet
                dsSOA.ReadXml(g_StructureFilename)
            End If

            Dim ProviderName As String = "System.Data.Odbc"
            Dim sQuery As String = ""
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            Dim dtADM1 As System.Data.DataTable
            Dim dtOADM As System.Data.DataTable
            Dim dtOCRD As System.Data.DataTable
            Dim dtOCTG As System.Data.DataTable
            Dim dtOSLP As System.Data.DataTable
            Dim dtSOC1 As System.Data.DataTable
            Dim dtSOC2 As System.Data.DataTable

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            '--------------------------------------------------------
            'NCM_SOC
            '--------------------------------------------------------
            sQuery = "  SELECT * FROM """ & oCompany.CompanyDB & """.""@NCM_SOC_AP"" "
            sQuery &= " WHERE ""USERNAME"" = '" & g_sAPSOARunningDate & oCompany.UserName & "' "
            sQuery &= " ORDER BY ""CARDCODE"" "
            dtSOC1 = dsSOA.Tables("@NCM_SOC_AP")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtSOC1)

            If dtSOC1.Rows.Count <= 0 Then
                SBO_Application.StatusBar.SetText("No data found.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                dbConn.Close()
                Return False
            End If

            '--------------------------------------------------------
            'OCRD (Customer)
            '--------------------------------------------------------
            sQuery = "SELECT  ""StreetNo"", ""ZipCode"", ""Address"", ""Block"", ""City"", ""County"",""CardCode"",""CardName"",""CntctPrsn"",""Fax"",""Phone1"",""GroupNum"",""SlpCode"",IFNULL(""U_SOA_Bldg"",'') AS ""U_SOA_Bldg"" FROM """ & oCompany.CompanyDB & """.""OCRD"" WHERE ""CardType"" = 'S' "
            dtOCRD = dsSOA.Tables("OCRD")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOCRD)

            '--------------------------------------------------------
            'OADM (Company Details)
            '--------------------------------------------------------
            sQuery = "SELECT ""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""RevOffice"",""Phone1"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            dtOADM = dsSOA.Tables("OADM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)

            '--------------------------------------------------------
            'ADM1 (Company Details)
            '--------------------------------------------------------
            sQuery = "SELECT ""Block"",""City"",""Country"",""County"",""ZipCode"",""Street"" FROM """ & oCompany.CompanyDB & """.""ADM1"" "
            dtADM1 = dsSOA.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)

            '--------------------------------------------------------
            'OCTG 
            '--------------------------------------------------------
            sQuery = "SELECT ""PymntGroup"",""GroupNum"" FROM """ & oCompany.CompanyDB & """.""OCTG"" "
            dtOCTG = dsSOA.Tables("OCTG")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOCTG)

            '--------------------------------------------------------
            'OSLP
            '--------------------------------------------------------
            sQuery = "SELECT ""SlpCode"",""SlpName"" FROM """ & oCompany.CompanyDB & """.""OSLP"" "
            dtOSLP = dsSOA.Tables("OSLP")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOSLP)

            '--------------------------------------------------------
            'NCM_SOC2
            '--------------------------------------------------------
            sQuery = "SELECT ""ID"", ""NOTES"" FROM """ & oCompany.CompanyDB & """.""@NCM_SOC2"" WHERE ""ID"" ='1' "
            dtSOC2 = dsSOA.Tables("@NCM_SOC2")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtSOC2)


            dbConn.Close()

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PrepareDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

#Region "Logic Function"
    Private Function Setup_Notes() As Boolean
        ' HANA
        Dim bSuccess As Boolean = False
        Dim sQuery As String = ""
        Dim sCurrSchema As String = ""
        Dim iCount As Integer = 0
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sQuery = " SELECT current_schema FROM DUMMY "
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                sCurrSchema = oRec.Fields.Item(0).Value
            End If

            If sCurrSchema.Trim <> "" Then
                sQuery = "  select Count(*) from sys.objects "
                sQuery &= " where ""SCHEMA_NAME"" = '" & sCurrSchema & "' "
                sQuery &= " AND ""OBJECT_TYPE"" = 'TABLE '"
                sQuery &= " AND ""OBJECT_NAME"" = '@NCM_SOC2' "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    iCount = oRec.Fields.Item(0).Value
                End If

                If iCount <= 0 Then
                    sQuery = " CREATE TABLE ""@NCM_SOC2"" "
                    sQuery &= " (ID         NVARCHAR(8)         NOT NULL,"
                    sQuery &= " Notes      NVARCHAR(2000)      NOT NULL,"
                    sQuery &= " Image    BLOB)"
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)

                    sQuery = " INSERT INTO """ & oCompany.CompanyDB & """.""@NCM_SOC2"" "
                    sQuery &= " VALUES ("
                    sQuery &= " '1',"
                    sQuery &= " 'Note:   Any payments received after end of the month will be shown in next month''s statement."
                    sQuery &= "          If you do not agree with the above statement, please inform us immediately.'"
                    sQuery &= " , NULL) "
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                Else
                    iCount = 0
                    sQuery = " Select Count(*) from """ & oCompany.CompanyDB & """.""@NCM_SOC2"" WHERE ""ID"" = '1' "
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If oRec.RecordCount > 0 Then
                        iCount = Convert.ToInt32(oRec.Fields.Item(0).Value)
                    End If

                    If iCount <= 0 Then
                        sQuery = " INSERT INTO """ & oCompany.CompanyDB & """.""@NCM_SOC2"" "
                        sQuery &= " VALUES ("
                        sQuery &= " '1',"
                        sQuery &= " 'Note:   Any payments received after end of the month will be shown in next month''s statement."
                        sQuery &= "          If you do not agree with the above statement, please inform us immediately.'"
                        sQuery &= " , NULL) "
                        oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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
    Private Function Retrieve_Notes() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""NOTES"" FROM """ & oCompany.CompanyDB & """.""@NCM_SOC2"" WHERE ""ID"" ='1'")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                oForm.DataSources.UserDataSources.Item("Notes").ValueEx = oRec.Fields.Item(0).Value
            End If
            oRec = Nothing
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[APSOA].[Retrieve_Notes]:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function Save_Settings() As Boolean
        Dim Notes As String = ""
        Dim BitmapPath As String = ""
        Dim ImagePath As String = ""
        Dim Image As Byte()
        Dim FileStrm As FileStream
        Dim BinReader As BinaryReader
        Dim sQuery As String = ""
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            ShowStatus("Status: Saving Settings...")
            Notes = oForm.DataSources.UserDataSources.Item("Notes").ValueEx
            Notes = Notes.Replace("'", "''")

            sQuery = "SELECT IFNULL(""BitmapPath"",'') FROM """ & oCompany.CompanyDB & """.""OADP"" "
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                BitmapPath = oRec.Fields.Item(0).Value
            End If

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
            'Read the file 
            If ImagePath.Trim <> "" Then
                FileStrm = New FileStream(ImagePath, FileMode.Open)
                BinReader = New BinaryReader(FileStrm)
                Image = BinReader.ReadBytes(BinReader.BaseStream.Length)
                FileStrm.Close()
                BinReader.Close()

                'sQuery = "UPDATE [@NCM_SOC2] SET Notes='" & Notes & "', Image=@Image WHERE ID = '1'"
                'cmd = New SqlCommand(sQuery, SQLDbConnection)
                'cmd.Parameters.Add("@Image", Image)
                'cmd.ExecuteNonQuery()
            Else
                'sQuery = "UPDATE [@NCM_SOC2] SET Notes='" & Notes & "', Image=0x0 WHERE ID = '1'"
                'cmd = New SqlCommand(sQuery, SQLDbConnection)
                'cmd.ExecuteNonQuery()
            End If

            sQuery = "UPDATE """ & oCompany.CompanyDB & """.""@NCM_SOC2"" SET ""NOTES"" ='" & Notes & "' WHERE ""ID"" = '1'"
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[APSOA].[Save_Settings]:" & ex.Message)
            Return False
        End Try
    End Function
    Private Function ExecuteProcedure() As Boolean
        Dim sDate As String = ""
        Dim sBBF As String = "N"
        Dim bSuccess As Boolean = False
        Dim iRowsAffected As Integer = 0
        Dim sQuery As String = ""
        Dim sBPCodeFr As String = ""
        Dim sBPCodeTo As String = ""
        Dim sBPGrpFr As String = ""
        Dim sBPGrpTo As String = ""
        Dim sSlsFr As String = ""
        Dim sSlsTo As String = ""
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
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
            oEdit = oForm.Items.Item("etBPCode").Specific
            ' BPCode = CType(IIf(oEdit.Value = "", "%", "%" & oEdit.Value.Replace("*", "%")) & "%", String).Trim
            BPCode = oEdit.Value

            'Get AsAtDate, FromDate
            oEdit = oForm.Items.Item("etDateAsAt").Specific
            sDate = oEdit.Value.Trim
            If sDate = "" Then Throw New Exception("Error: As At Date is empty!")
            AsAtDate = New DateTime(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
            If (oForm.DataSources.UserDataSources.Item("txtDateFr").ValueEx.Length = 0) Then
                Throw New Exception("Error: Date From is empty!")
            Else
                FromDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("txtDateFr").ValueEx, "yyyyMMdd", Nothing)
            End If

            'Get IsBBF
            oCheck = oForm.Items.Item("ckBBF").Specific
            If oCheck.Checked Then IsBBF = "Y" Else IsBBF = "N"

            'Get IsGAT
            oCheck = oForm.Items.Item("ckGAT").Specific
            If oCheck.Checked Then IsGAT = "Y" Else IsGAT = "N"

            'Get IsGAT
            oCheck = oForm.Items.Item("ckGAT").Specific
            If oCheck.Checked Then IsGAT = "Y" Else IsGAT = "N"

            'Set the query
            sQuery = " CALL SP_SOA_AP ('"
            sQuery &= g_sAPSOARunningDate & oCompany.UserName & "','"
            sQuery &= sBPCodeFr.Replace("'", "''") & "','"
            sQuery &= sBPCodeTo.Replace("'", "''") & "','"
            sQuery &= sBPGrpFr & "','"
            sQuery &= sBPGrpTo & "','"
            sQuery &= sSlsFr & "','"
            sQuery &= sSlsTo & "','"
            sQuery &= BPCode.Replace("'", "''") & "','"
            sQuery &= FromDate.ToString("yyyyMMdd") & "','"
            sQuery &= AsAtDate.ToString("yyyyMMdd") & "','"
            sQuery &= IsBBF & "','"
            sQuery &= IsGAT & "',"

            oCheck = oForm.Items.Item("ckExc").Specific
            'HANA
            'If oCheck.Checked Then sQuery &= "1" Else sQuery &= "0"
            If oCheck.Checked Then sQuery &= "1)" Else sQuery &= "0)"

            Try
                ShowStatus("Status: Executing Procedure...")
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                oRec = Nothing

                ShowStatus("Status: Completed!")
                bSuccess = True
            Catch ex As Exception
                bSuccess = False
                Throw ex
            End Try

            SBO_Application.StatusBar.SetText("Completed Successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.MessageBox("[APSOA].[ExecuteProcedure]:" & ex.ToString)
        End Try
        Return bSuccess
    End Function
    Private Function ValidateParameter() As Boolean
        Try
            oForm.ActiveItem = "etBPCode"
            Dim oRecordsetLn As SAPbobsCOM.Recordset
            oRecordsetLn = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
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

            If (oForm.DataSources.UserDataSources.Item("DateAsAt").ValueEx.Length = 0) Then
                SBO_Application.StatusBar.SetText("Please enter As At date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "etDateAsAt"
                Return False
            End If

            If (oForm.DataSources.UserDataSources.Item("txtDateFr").ValueEx.Length = 0) Then
                SBO_Application.StatusBar.SetText("Please enter from date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "txtDateFr"
                Return False
            End If

            FromDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("txtDateFr").ValueEx, "yyyyMMdd", Nothing)
            AsAtDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("DateAsAt").ValueEx, "yyyyMMdd", Nothing)

            If (FromDate >= AsAtDate) Then
                SBO_Application.StatusBar.SetText("From Date must be less than As At Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "txtDateFr"
                Return False
            End If
            If (oForm.DataSources.UserDataSources.Item("DateAsAt").ValueEx.Length = 0) Then
                SBO_Application.StatusBar.SetText("Please enter as at date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "etDateAsAt"
                Return False
            End If

            SBO_Application.StatusBar.SetText(String.Empty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[APSOA].[ValidateParameter] - " & ex.Message, 1, "OK", String.Empty, String.Empty)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Friend Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                If pVal.FormUID = "ncmSOA_AP" Then
                    Select Case pVal.ItemUID
                        Case "ckHFN"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                oForm.Items.Item("etBPCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        Case "btnExecute"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                    Return ValidateParameter()
                                End If
                            End If
                    End Select
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

                If pVal.FormUID = "ncmSOA_AP" Then
                    Select Case pVal.ItemUID
                        Case "btnExecute"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                    Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                    myThread.SetApartmentState(Threading.ApartmentState.STA)
                                    myThread.Start()
                                End If
                            End If
                        Case "ckHFN"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                oForm.Items.Item("etNotes").Enabled = Not (oForm.Items.Item("etNotes").Enabled)
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[APSOA].[ItemEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class