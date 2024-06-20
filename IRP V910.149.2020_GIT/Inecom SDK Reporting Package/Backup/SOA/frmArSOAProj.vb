Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class frmArSOAProj

#Region "Global Variables"
    Private BPCode As String
    Private AsAtDate As DateTime
    Private FromDate As DateTime
    Private IsBBF As String
    Private IsGAT As String
    Private dsSOA As DataSet

    Private g_StructureFilename As String = ""
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Private Const ClientCompany As CompanyCode = CompanyCode.General
    Private Const EmbeddedType As Boolean = False

    Private oFormARSOA As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
    Private g_sARSOARunningDate As String = ""

#End Region

#Region "Intialize Application"
    Public Sub New()
        Try

        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        Dim oPictureBox As SAPbouiCOM.PictureBox
        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_ARSOA_PROJ.srf") Then
            oFormARSOA = SBO_Application.Forms.Item("NCM_ARSOA_PROJ")
            oPictureBox = oFormARSOA.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"

            If ClientCompany = CompanyCode.AE Then
                oFormARSOA.Items.Item("ckLogo").Visible = False
            End If
            oFormARSOA.Items.Item("lbStatus").FontSize = 10
            oFormARSOA.Items.Item("lbStyleOpt").TextStyle = 4

            SetDatasource()
            NotesSetup()
            RetrieveNotes()
            SetupChooseFromList()

            oFormARSOA.Items.Item("ckLayout").Visible = False
            oFormARSOA.DataSources.UserDataSources.Item("ckLayout").ValueEx = "N"
            oFormARSOA.Items.Item("txtBPFr").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oFormARSOA.Visible = True
        Else
            Try
                If oFormARSOA.Visible = False Then
                    oFormARSOA.Close()
                Else
                    oFormARSOA.Select()
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub
    Private Sub SetDatasource()
        Try
            With oFormARSOA.DataSources.UserDataSources
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
                .Add("ckLayout", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("cbBased", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

                .Add("Notes", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1500)
                .Add("txtBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
                .Add("txtBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
                .Add("txtPrjFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
                .Add("txtPrjTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
                .Add("txtBPGFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtBPGTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtSlsFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtSlsTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            End With

            oEdit = oFormARSOA.Items.Item("etBPCode").Specific
            oEdit.DataBind.SetBound(True, "", "BPCode")
            oEdit = oFormARSOA.Items.Item("etDateAsAt").Specific
            oEdit.DataBind.SetBound(True, "", "DateAsAt")
            oEdit.Value = Now.ToString("yyyyMMdd")
            oEdit = oFormARSOA.Items.Item("etNotes").Specific
            oEdit.DataBind.SetBound(True, "", "Notes")

            oEdit = oFormARSOA.Items.Item("txtDateFr").Specific
            oEdit.DataBind.SetBound(True, String.Empty, "txtDateFr")
            oEdit.Value = Now.ToString("yyyyMM01")

            oEdit = oFormARSOA.Items.Item("txtBPFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPFr")
            oEdit = oFormARSOA.Items.Item("txtBPTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPTo")
            oEdit = oFormARSOA.Items.Item("txtPrjFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtPrjFr")
            oEdit = oFormARSOA.Items.Item("txtPrjTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtPrjTo")
            oEdit = oFormARSOA.Items.Item("txtBPGFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPGFr")
            oEdit = oFormARSOA.Items.Item("txtBPGTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPGTo")
            oEdit = oFormARSOA.Items.Item("txtSlsFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtSlsFr")
            oEdit = oFormARSOA.Items.Item("txtSlsTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtSlsTo")

            oCombo = oFormARSOA.Items.Item("cbDateType").Specific
            oCombo.ValidValues.Add("0", "Document Date")
            oCombo.ValidValues.Add("1", "Due Date")
            oCombo.ValidValues.Add("2", "Posting Date")
            oCombo.DataBind.SetBound(True, "", "DateType")
            oFormARSOA.DataSources.UserDataSources.Item("DateType").ValueEx = "0"

            oCombo = oFormARSOA.Items.Item("cbBased").Specific
            oCombo.ValidValues.Add("0", "Posting Date")
            oCombo.ValidValues.Add("1", "Document Date")
            oCombo.DataBind.SetBound(True, "", "cbBased")
            oFormARSOA.DataSources.UserDataSources.Item("cbBased").ValueEx = "0"

            oCombo = oFormARSOA.Items.Item("cbPrdType").Specific
            oCombo.ValidValues.Add("0", "Every 30 Days")
            oCombo.ValidValues.Add("1", "Every Month")
            oCombo.DataBind.SetBound(True, "", "Period")
            oFormARSOA.DataSources.UserDataSources.Item("Period").ValueEx = "0"

            oCheck = oFormARSOA.Items.Item("ckLogo").Specific
            oCheck.DataBind.SetBound(True, "", "Logo")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            If ClientCompany = CompanyCode.AMS Then
                oFormARSOA.Items.Item("ckLogo").Enabled = False
            End If

            oCheck = oFormARSOA.Items.Item("ckHDR").Specific
            oCheck.DataBind.SetBound(True, "", "HDR")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
            oCheck.DataBind.SetBound(True, "", "BBF")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckSNP").Specific
            oCheck.DataBind.SetBound(True, "", "SNP")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
            oCheck.DataBind.SetBound(True, "", "GAT")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckHAS").Specific
            oCheck.DataBind.SetBound(True, "", "HAS")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckHFN").Specific
            oCheck.DataBind.SetBound(True, "", "HFN")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckExc").Specific
            oCheck.DataBind.SetBound(True, "", "EXC")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckLayout").Specific
            oCheck.DataBind.SetBound(True, "", "ckLayout")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"


        Catch ex As Exception
            SBO_Application.MessageBox("[SetDatasource] : " & ex.Message)
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
            oCFLs = oFormARSOA.ChooseFromLists

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

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtBPFr").Specific, SAPbouiCOM.EditText)
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

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtBPTo").Specific, SAPbouiCOM.EditText)
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

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtBPGFr").Specific, SAPbouiCOM.EditText)
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

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtBPGTo").Specific, SAPbouiCOM.EditText)
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

            oEditLn = DirectCast(oFormARSOA.Items.Item("etBPCode").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_BPCode"
            oEditLn.ChooseFromListAlias = "CardCode"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 53
            oCFLCreation.UniqueID = "CFL_SPFrom"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtSlsFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_SPFrom"
            oEditLn.ChooseFromListAlias = "SlpName"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 53
            oCFLCreation.UniqueID = "CFL_SPTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtSlsTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_SPTo"
            oEditLn.ChooseFromListAlias = "SlpName"
            ' ----------------------------------------

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ProjectCodes
            oCFLCreation.UniqueID = "CFL_PRJFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtPrjFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_PRJFr"
            oEditLn.ChooseFromListAlias = "PrjCode"
            ' ----------------------------------------

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ProjectCodes
            oCFLCreation.UniqueID = "CFL_PRJTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtPrjTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_PRJTo"
            oEditLn.ChooseFromListAlias = "PrjCode"
            ' ----------------------------------------
        Catch ex As Exception
            Throw New Exception("[ARSOA].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub ShowStatus(ByVal sStatus As String)
        Try
            Dim oStaticText As SAPbouiCOM.StaticText = oFormARSOA.Items.Item("lbStatus").Specific
            oStaticText.Caption = sStatus
        Catch ex As Exception
            SBO_Application.MessageBox("[ShowStatus] : " & ex.Message)
        End Try
    End Sub
    Private Sub LoadViewer()
        oFormARSOA.Items.Item("btnExecute").Enabled = False
        Try
            Dim sAttachPath As String = ""
            Dim frm As New Hydac_FormViewer
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), IFNULL(""AttachPath"",'')  FROM ""OADP"" ")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                g_sARSOARunningDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
                sAttachPath = Convert.ToString(oRec.Fields.Item(1).Value).Trim
            End If

            Dim oWeb As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sWeb As String = ""
            Dim bWeb As Boolean = False

            'g_sARSOARunningDate = "2015-02-24 17:03:12.6400"

            frm.DatabaseServer = oCompany.Server
            frm.DatabaseName = oCompany.CompanyDB

            Dim bIsContinue As Boolean = False
            Try
                If SaveSettings() Then
                    If ExecuteProcedure() Then
                        If PrepareDataset() Then
                            ' =================================================================
                            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim iCount As Integer = -1

                            ' IF USE EMAIL TO SEND SOA
                            If IsIncludeModule(ReportName.SOA_Email_Config) Then
                                iCount = 0
                                iCount = SBO_Application.MessageBox("Please select your option." & vbNewLine & "1. Click ""Yes"" to send email." & vbNewLine & "2. Click ""No"" to preview only.", 1, "Yes", "No", String.Empty)
                                If iCount = 1 Then
                                    Dim ds As New dsEmail()
                                    Dim al As New System.Collections.ArrayList()
                                    Dim sOutput As String = String.Empty
                                    Dim sTempDirectory As String = ""

                                    If bWeb Then
                                        sTempDirectory = sAttachPath.Trim
                                        If sTempDirectory.Substring(sTempDirectory.Length - 1, 1) = "\" Then
                                            sTempDirectory = sTempDirectory.Substring(0, sTempDirectory.Length - 1)
                                        End If
                                    Else
                                        sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\AR_SOA\" & oCompany.CompanyDB
                                    End If

                                    Dim sPathFormat As String = "{0}\{1}_SOA_{2}.pdf"
                                    sQuery = "  SELECT DISTINCT ""CARDCODE"" "
                                    sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
                                    sQuery &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                                    sQuery &= " ORDER BY ""CARDCODE"", ""PROJECT"" "
                                    oRecord.DoQuery(sQuery)
                                    If (oRecord.RecordCount > 0) Then
                                        oRecord.MoveFirst()
                                        Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                                        If Not di.Exists Then
                                            di.Create()
                                        End If

                                        While Not oRecord.EoF
                                            sOutput = oRecord.Fields.Item(0).Value

                                            al.Add(sOutput)
                                            oCombo = oFormARSOA.Items.Item("cbDateType").Specific
                                            If oCombo.Selected Is Nothing Then
                                                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            End If
                                            frm.Report = oCombo.Selected.Value

                                            oCombo = oFormARSOA.Items.Item("cbPrdType").Specific
                                            If oCombo.Selected Is Nothing Then
                                                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            End If
                                            frm.Period = oCombo.Selected.Value

                                            oCheck = oFormARSOA.Items.Item("ckLogo").Specific
                                            frm.HideLogo = IIf(oCheck.Checked, True, False)
                                            oCheck = oFormARSOA.Items.Item("ckHDR").Specific
                                            frm.HideHeader = IIf(oCheck.Checked, True, False)
                                            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
                                            frm.IsBBF = IIf(oCheck.Checked, 1, 0)
                                            oCheck = oFormARSOA.Items.Item("ckSNP").Specific
                                            frm.IsSNP = IIf(oCheck.Checked, 1, 0)
                                            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
                                            frm.IsGAT = IIf(oCheck.Checked, 1, 0)
                                            oCheck = oFormARSOA.Items.Item("ckHAS").Specific
                                            frm.IsHAS = IIf(oCheck.Checked, 1, 0)
                                            oCheck = oFormARSOA.Items.Item("ckHFN").Specific
                                            frm.IsHFN = IIf(oCheck.Checked, 1, 0)

                                            frm.Dataset = dsSOA
                                            frm.IsShared = g_bIsShared
                                            frm.SharedReportName = g_sReportFilename
                                            frm.ARSOARunningDate = g_sARSOARunningDate & oCompany.UserName
                                            frm.DBUsernameViewer = DBUsername
                                            frm.DBPasswordViewer = DBPassword
                                            frm.Username = oCompany.UserName
                                            frm.AsAtDate = AsAtDate.ToString("yyyyMMdd")
                                            frm.ReportName = ReportName.ARSOA_BY_PROJECT
                                            frm.CompanySOA = ClientCompany
                                            frm.Landscape = "N"
                                            frm.IsExport = True
                                            frm.ExportCustomerCode = sOutput
                                            frm.CrystalReportExportType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                                            frm.CrystalReportExportPath = String.Format(sPathFormat, di.FullName, sOutput, AsAtDate.ToString("ddMMyyyy"))
                                            frm.OPEN_HANADS_ARSOA_PROJ()

                                            oRecord.MoveNext()
                                        End While

                                        Dim dr As dsEmail.PreviewDTRow

                                        sQuery = "  SELECT T0.""CARDCODE"", T0.""DOCCUR"", "
                                        sQuery &= " SUM(T0.""DOCTOTALFC"" - T0.""CLOSEPAID"") ""Balance"", "
                                        sQuery &= " IFNULL(T1.""U_SOA_MailTo"",'') ""Email"", "
                                        sQuery &= " IFNULL(T1.""CardName"",'') ""CardName"" "
                                        sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" T0 "
                                        sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T1 "
                                        sQuery &= " ON T0.""CARDCODE"" = T1.""CardCode"" "
                                        sQuery &= " WHERE T0.""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                                        sQuery &= " GROUP BY T0.""CARDCODE"", T1.""CardName"", T0.""DOCCUR"", T1.""U_SOA_MailTo"" "
                                        sQuery &= " ORDER BY T0.""CARDCODE"", T1.""CardName"", T0.""PROJECT"", T0.""DOCCUR"" "

                                        oRecord.DoQuery(sQuery)
                                        If oRecord.RecordCount > 0 Then
                                            oRecord.MoveFirst()
                                            While Not oRecord.EoF
                                                dr = ds.PreviewDT.NewPreviewDTRow()
                                                dr.Attachment = String.Format(sPathFormat, di.FullName, oRecord.Fields.Item("CardCode").Value, AsAtDate.ToString("ddMMyyyy"))
                                                dr.Balance = oRecord.Fields.Item("Balance").Value
                                                dr.CardCode = oRecord.Fields.Item("CardCode").Value
                                                dr.CardName = oRecord.Fields.Item("CardName").Value
                                                dr.Currency = oRecord.Fields.Item("DocCur").Value
                                                dr.EmailTo = oRecord.Fields.Item("Email").Value
                                                dr.IsEmail = IIf(dr.Balance > 0, 1, 0)

                                                dr.Table.Rows.Add(dr)
                                                oRecord.MoveNext()
                                            End While
                                        End If

                                        SubMain.oFrmSendEmailProj.ReportName = ReportName.ARSOA_BY_PROJECT
                                        SubMain.oFrmSendEmailProj.StatementAsAtDate = AsAtDate
                                        SubMain.oFrmSendEmailProj.StatementDataTable = ds.PreviewDT
                                        SubMain.oFrmSendEmailProj.LoadForm()
                                        Hydac_FormViewer.Close()
                                        Return
                                    End If

                                End If
                            End If
                            ' END - IF USE EMAIL TO SEND SOA

                            oCombo = oFormARSOA.Items.Item("cbDateType").Specific
                            If oCombo.Selected Is Nothing Then
                                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                            frm.Report = oCombo.Selected.Value

                            oCombo = oFormARSOA.Items.Item("cbPrdType").Specific
                            If oCombo.Selected Is Nothing Then
                                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                            frm.Period = oCombo.Selected.Value
                            frm.Dataset = dsSOA
                            frm.IsShared = g_bIsShared
                            frm.SharedReportName = g_sReportFilename
                            frm.DBUsernameViewer = DBUsername
                            frm.DBPasswordViewer = DBPassword
                            frm.Username = oCompany.UserName
                            frm.AsAtDate = AsAtDate.ToString("yyyyMMdd")
                            frm.ReportName = ReportName.ARSOA_BY_PROJECT
                            frm.CompanySOA = ClientCompany
                            frm.Landscape = "N"
                            frm.ARSOARunningDate = g_sARSOARunningDate & oCompany.UserName

                            oCheck = oFormARSOA.Items.Item("ckLogo").Specific
                            frm.HideLogo = IIf(oCheck.Checked, True, False)
                            oCheck = oFormARSOA.Items.Item("ckHDR").Specific
                            frm.HideHeader = IIf(oCheck.Checked, True, False)
                            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
                            frm.IsBBF = IIf(oCheck.Checked, 1, 0)
                            oCheck = oFormARSOA.Items.Item("ckSNP").Specific
                            frm.IsSNP = IIf(oCheck.Checked, 1, 0)
                            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
                            frm.IsGAT = IIf(oCheck.Checked, 1, 0)
                            oCheck = oFormARSOA.Items.Item("ckHAS").Specific
                            frm.IsHAS = IIf(oCheck.Checked, 1, 0)
                            oCheck = oFormARSOA.Items.Item("ckHFN").Specific
                            frm.IsHFN = IIf(oCheck.Checked, 1, 0)

                            bIsContinue = True
                            oRecord = Nothing

                            ' =================================================================
                        End If
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                oFormARSOA.Items.Item("btnExecute").Enabled = True
            End Try
            If bIsContinue Then
                frm.ShowDialog()
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[LoadViewer] : " & ex.Message)
        End Try
    End Sub
    Private Function IsSharedFileExist() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            g_sReportFilename = ""
            g_StructureFilename = ""

            sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.ARSOA_BY_PROJECT) & "'"

            g_sReportFilename = GetSharedFilePath(ReportName.ARSOA_BY_PROJECT)
            If g_sReportFilename.Trim <> "" Then
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
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[ARSOA].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
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
            Dim dtOPRJ As System.Data.DataTable
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
            sQuery = "  SELECT * FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
            sQuery &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "' "
            dtSOC1 = dsSOA.Tables("@NCM_SOC")
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
            sQuery = "SELECT ""StreetNo"", ""ZipCode"", ""Phone2"", ""Address"", ""Block"", ""City"", ""County"",""CardCode"",""CardName"",""CntctPrsn"",""Fax"",""Phone1"",""GroupNum"",""SlpCode"",IFNULL(""U_SOA_Bldg"",'') AS ""U_SOA_Bldg"" FROM """ & oCompany.CompanyDB & """.""OCRD"" WHERE ""CardType"" = 'C' "
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
            'OPRJ
            '--------------------------------------------------------
            sQuery = "SELECT ""PrjCode"",""PrjName"" FROM """ & oCompany.CompanyDB & """.""OPRJ"" "
            dtOPRJ = dsSOA.Tables("OPRJ")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOPRJ)

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
    Private Function NotesSetup() As Boolean
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
                    sQuery &= " NOTES      NVARCHAR(2000)      NOT NULL,"
                    sQuery &= " IMAGE    BLOB)"
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)

                    sQuery = " INSERT INTO ""@NCM_SOC2"" "
                    sQuery &= " VALUES ("
                    sQuery &= " '1',"
                    sQuery &= " 'Note:   Any payments received after end of the month will be shown in next month''s statement."
                    sQuery &= "          If you do not agree with the above statement, please inform us immediately.'"
                    sQuery &= " , NULL) "
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                Else
                    iCount = 0
                    sQuery = " Select Count(*) from ""@NCM_SOC2"" WHERE ""ID"" = '1' "
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If oRec.RecordCount > 0 Then
                        iCount = Convert.ToInt32(oRec.Fields.Item(0).Value)
                    End If

                    If iCount <= 0 Then
                        sQuery = "  INSERT INTO ""@NCM_SOC2"" "
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
    Private Sub RetrieveNotes()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""NOTES"" FROM ""@NCM_SOC2"" WHERE ""ID"" ='1'")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                oFormARSOA.DataSources.UserDataSources.Item("Notes").ValueEx = oRec.Fields.Item(0).Value
            End If
            oRec = Nothing
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[RetrieveNotes] : " & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function SaveSettings() As Boolean
        Dim Notes As String = ""
        Dim BitmapPath As String = ""
        Dim ImagePath As String = ""
        Dim Image As Byte()
        Dim sQuery As String
        Dim FileStrm As FileStream
        Dim BinReader As BinaryReader
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            ShowStatus("Status: Saving Settings...")
            Notes = oFormARSOA.DataSources.UserDataSources.Item("Notes").ValueEx
            Notes = Notes.Replace("'", "''")
            sQuery = "SELECT IFNULL(""BitmapPath"",'') FROM """ & oCompany.CompanyDB & """.""OADP"" "
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

            'Read the file 
            If ImagePath.Trim <> "" Then
                FileStrm = New FileStream(ImagePath, FileMode.Open)
                BinReader = New BinaryReader(FileStrm)
                Image = BinReader.ReadBytes(BinReader.BaseStream.Length)
                FileStrm.Close()
                BinReader.Close()
            End If

            sQuery = "UPDATE """ & oCompany.CompanyDB & """.""@NCM_SOC2"" SET ""NOTES"" ='" & Notes & "' WHERE ""ID"" = '1'"
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[SaveImages] : " & ex.Message)
            Return False
        End Try
    End Function
    Private Function ExecuteProcedure() As Boolean
        Dim sDate As String = String.Empty
        Dim sBBF As String = "N"
        Dim bSuccess As Boolean = False
        Dim iRowsAffected As Integer = 0
        Dim sQuery As String = String.Empty
        Dim sBPCodeFr As String = String.Empty
        Dim sBPCodeTo As String = String.Empty
        Dim sBPGrpFr As String = String.Empty
        Dim sBPGrpTo As String = String.Empty
        Dim sSlsFr As String = String.Empty
        Dim sSlsTo As String = String.Empty
        Dim sAsAtDate As String = ""
        Dim sFromDate As String = ""
        Dim sIsExc As String = "0"
        Dim sBasedOn As String = "0"
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sPrjFr As String = ""
        Dim sPrjTo As String = ""

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

            'Return True
            sBasedOn = oFormARSOA.DataSources.UserDataSources.Item("cbBased").ValueEx

            'Get Parameter Value
            oEdit = oFormARSOA.Items.Item("txtPrjFr").Specific
            sPrjFr = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("txtPrjTo").Specific
            sPrjTo = oEdit.Value.ToString.Trim.Replace("'", "''")

            oEdit = oFormARSOA.Items.Item("txtBPFr").Specific
            sBPCodeFr = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("txtBPTo").Specific
            sBPCodeTo = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("txtBPGFr").Specific
            sBPGrpFr = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("txtBPGTo").Specific
            sBPGrpTo = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("txtSlsFr").Specific
            sSlsFr = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("txtSlsTo").Specific
            sSlsTo = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("etBPCode").Specific
            BPCode = oEdit.Value.ToString.Trim.Replace("'", "''")
            oEdit = oFormARSOA.Items.Item("etDateAsAt").Specific
            sDate = oEdit.Value.ToString.Trim

            'Get IsBBF
            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
            If oCheck.Checked Then IsBBF = "Y" Else IsBBF = "N"

            'Get IsGAT
            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
            If oCheck.Checked Then IsGAT = "Y" Else IsGAT = "N"

            'GEt IsEXC
            oCheck = oFormARSOA.Items.Item("ckExc").Specific
            If oCheck.Checked Then sIsExc = "1" Else sIsExc = "0"

            If sDate = "" Then Throw New Exception("Error: As At Date is empty!")
            AsAtDate = New DateTime(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
            sAsAtDate = oFormARSOA.DataSources.UserDataSources.Item("DateAsAt").ValueEx

             If (oFormARSOA.DataSources.UserDataSources.Item("txtDateFr").ValueEx.Length = 0) Then
                Throw New Exception("Error: From date is empty!")
            Else
                FromDate = DateTime.ParseExact(oFormARSOA.DataSources.UserDataSources.Item("txtDateFr").ValueEx, "yyyyMMdd", Nothing)
                sFromDate = oFormARSOA.DataSources.UserDataSources.Item("txtDateFr").ValueEx
            End If

            'Set the query
            sQuery = "CALL SP_SOA ('"
            sQuery &= g_sARSOARunningDate & oCompany.UserName & "','"
            sQuery &= sBPCodeFr.Replace("'", "''") & "','"
            sQuery &= sBPCodeTo.Replace("'", "''") & "','"
            sQuery &= sBPGrpFr & "','"
            sQuery &= sBPGrpTo & "','"
            sQuery &= sSlsFr & "','"
            sQuery &= sSlsTo & "','"
            sQuery &= BPCode.Replace("'", "''") & "','"
            sQuery &= sFromDate & "','"
            sQuery &= sAsAtDate & "','"
            sQuery &= IsBBF & "','"
            sQuery &= IsGAT & "','"
            sQuery &= sIsExc & "','"
            sQuery &= sBasedOn & "')"

            Try
                ShowStatus("Status: Executing Procedure...")
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                oRec = Nothing

                ShowStatus("Status: Completed!")
                bSuccess = True


                ' if project code range
                ' if blank ==> all ==> no action
                Dim oDelete As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim sDelete As String = ""

                sDelete = "  UPDATE """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
                sDelete &= " SET ""PROJECT"" = ''"
                sDelete &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                sDelete &= " AND IFNULL(""PROJECT"",'') = '' "
                oDelete = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oDelete.DoQuery(sDelete)

                Select Case sPrjFr
                    Case ""
                        Select Case sPrjTo
                            Case ""
                                 'do nothing
                            Case Else
                                ' if there is projTo ==> delete all after projTo
                                sDelete = "  DELETE FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
                                sDelete &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                                sDelete &= " AND IFNULL(""PROJECT"",'') <> '' "
                                sDelete &= " AND IFNULL(""PROJECT"",'') > '" & sPrjTo & "' "
                        End Select
                    Case Else
                        Select Case sPrjTo
                            Case ""
                                ' if there is projFrom ==> delete all before projFrom
                                sDelete = "  DELETE FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
                                sDelete &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                                sDelete &= " AND IFNULL(""PROJECT"",'') <> '' "
                                sDelete &= " AND IFNULL(""PROJECT"",'') < '" & sPrjFr & "' "
                            Case Else
                                ' if there are projFrom and projTo
                                sDelete = "  DELETE FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
                                sDelete &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                                sDelete &= " AND IFNULL(""PROJECT"",'') NOT BETWEEN '" & sPrjFr & "' AND '" & sPrjTo & "' "
                        End Select
                End Select

                If sDelete <> "" Then
                    oDelete = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelete.DoQuery(sDelete)

                    sDelete = "  SELECT * FROM """ & oCompany.CompanyDB & """.""@NCM_SOC"" "
                    sDelete &= " WHERE ""USERNAME"" = '" & g_sARSOARunningDate & oCompany.UserName & "'"
                    oDelete = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oDelete.DoQuery(sDelete)
                    If oDelete.RecordCount <= 0 Then
                        bSuccess = False
                        SBO_Application.StatusBar.SetText("No records found based on the input parameters.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    End If
                End If
                oDelete = Nothing

            Catch ex As Exception
                bSuccess = False
                Throw ex
            End Try
            SBO_Application.StatusBar.SetText("Completed Successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[ExecuteProcedure]" & vbNewLine & ex.Message)
        End Try
        Return bSuccess
    End Function
    Private Function ValidateParameter() As Boolean
        Try
            oFormARSOA.ActiveItem = "etBPCode"
            Dim oRec As SAPbobsCOM.Recordset = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty
            Dim sQuery As String = String.Empty

            With oFormARSOA.DataSources.UserDataSources
                sStart = .Item("txtBPFr").ValueEx.Trim.Replace("'", "''")
                sEnd = .Item("txtBPTo").ValueEx.Trim.Replace("'", "''")
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("BP Code From is greater than BP Code To.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtBPFr"
                        Return False
                    End If
                End If

                sStart = .Item("txtPrjFr").ValueEx.Trim.Replace("'", "''")
                sEnd = .Item("txtPrjTo").ValueEx.Trim.Replace("'", "''")
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("Project Code From is greater than Project Code To.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtPrjFr"
                        Return False
                    End If
                End If

                sStart = .Item("txtBPGFr").ValueEx.Trim.Replace("'", "''")
                sEnd = .Item("txtBPGTo").ValueEx.Trim.Replace("'", "''")
                If (sStart.Length > 0) Then
                    sQuery = "SELECT ""GroupCode"" FROM ""OCRG"" WHERE ""GroupType"" = 'C' AND ""GroupCode"" = '" & sStart & "'"
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If (oRec.RecordCount = 0) Then
                        SBO_Application.StatusBar.SetText("Invalid BP Group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtBPGFr"
                        Return False
                    End If
                End If
                If (sEnd.Length > 0) Then
                    sQuery = "SELECT ""GroupCode"" FROM ""OCRG"" WHERE ""GroupType"" = 'C' AND ""GroupCode"" = '" & sEnd & "'"
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If (oRec.RecordCount = 0) Then
                        SBO_Application.StatusBar.SetText("Invalid BP Group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtBPGTo"
                        Return False
                    End If
                End If
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("BP Group from is greater than BP Group to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtBPGFr"
                        Return False
                    End If
                End If

                sStart = .Item("txtSlsFr").ValueEx.Trim.Replace("'", "''")
                sEnd = .Item("txtSlsTo").ValueEx.Trim.Replace("'", "''")
                If (sStart.Length > 0) Then
                    sQuery = "SELECT ""SlpName"" FROM ""OSLP"" WHERE ""SlpName"" = '" & sStart & "'"
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If (oRec.RecordCount = 0) Then
                        SBO_Application.StatusBar.SetText("Invalid Sales Employee", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtSlsFr"
                        Return False
                    End If
                End If
                If (sEnd.Length > 0) Then
                    sQuery = "SELECT ""SlpName"" FROM ""OSLP"" WHERE ""SlpName"" = '" & sEnd & "'"
                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If (oRec.RecordCount = 0) Then
                        SBO_Application.StatusBar.SetText("Invalid Sales Employee", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtSlsTo"
                        Return False
                    End If
                End If
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("Sales Employee from is greater than Sales Employee to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oFormARSOA.ActiveItem = "txtSlsFr"
                        Return False
                    End If
                End If

                If (.Item("txtDateFr").ValueEx.Length = 0) Then
                    SBO_Application.StatusBar.SetText("Please enter from date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtDateFr"
                    Return False
                End If

                If (.Item("DateAsAt").ValueEx.Length = 0) Then
                    SBO_Application.StatusBar.SetText("Please enter as at date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "etDateAsAt"
                    Return False
                End If

                AsAtDate = DateTime.ParseExact(.Item("DateAsAt").ValueEx, "yyyyMMdd", Nothing)
                FromDate = DateTime.ParseExact(.Item("txtDateFr").ValueEx, "yyyyMMdd", Nothing)

                If (FromDate >= AsAtDate) Then
                    SBO_Application.StatusBar.SetText("from date must be less than as at date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtDateFr"
                    Return False
                End If
            End With

            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Public Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "ckHFN"
                                oFormARSOA.Items.Item("etBPCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Case "btnExecute"
                                If (oFormARSOA.Items.Item(pVal.ItemUID).Enabled) Then
                                    Return ValidateParameter()
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
                            Dim sTemp As String = ""
                            With oFormARSOA.DataSources.UserDataSources
                                Select Case oCFLEvent.ChooseFromListUID
                                    Case "CFL_PRJFr"
                                        sTemp = oDataTable.GetValue("PrjCode", 0)
                                        .Item("txtPrjFr").ValueEx = sTemp
                                        Exit Select
                                    Case "CFL_PRJTo"
                                        sTemp = oDataTable.GetValue("PrjCode", 0)
                                        .Item("txtPrjTo").ValueEx = sTemp
                                        Exit Select
                                    Case "cflBPFr"
                                        sTemp = oDataTable.GetValue("CardCode", 0)
                                        .Item("txtBPFr").ValueEx = sTemp
                                        Exit Select
                                    Case "cflBPTo"
                                        sTemp = oDataTable.GetValue("CardCode", 0)
                                        .Item("txtBPTo").ValueEx = sTemp
                                        Exit Select
                                    Case "CFL_BPCode"
                                        sTemp = oDataTable.GetValue("CardCode", 0)
                                        .Item("BPCode").ValueEx = sTemp
                                        Exit Select
                                    Case "CFL_BGFrom"
                                        sTemp = oDataTable.GetValue("GroupCode", 0)
                                        .Item("txtBPGFr").ValueEx = sTemp
                                        Exit Select
                                    Case "CFL_BGTo"
                                        sTemp = oDataTable.GetValue("GroupCode", 0)
                                        .Item("txtBPGTo").ValueEx = sTemp
                                        Exit Select
                                    Case "CFL_SPFrom"
                                        sTemp = oDataTable.GetValue("SlpName", 0)
                                        .Item("txtSlsFr").ValueEx = sTemp
                                        Exit Select
                                    Case "CFL_SPTo"
                                        sTemp = oDataTable.GetValue("SlpName", 0)
                                        .Item("txtSlsTo").ValueEx = sTemp
                                        Exit Select
                                    Case Else
                                        Exit Select
                                End Select
                            End With
                            Return True
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "btnExecute"
                                If oFormARSOA.Items.Item(pVal.ItemUID).Enabled Then
                                    Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                    myThread.SetApartmentState(Threading.ApartmentState.STA)
                                    myThread.Start()
                                End If
                            Case "ckHFN"
                                oFormARSOA.Items.Item("etNotes").Enabled = Not (oFormARSOA.Items.Item("etNotes").Enabled)
                        End Select
                End Select
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[ARSOA].[ItemEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class