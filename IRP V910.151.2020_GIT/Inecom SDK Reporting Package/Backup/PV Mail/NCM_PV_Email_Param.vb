Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class NCM_PV_Email_Param

#Region "Global Variables"
    Private sQuery As String = ""
    Private BPCode As String
    Private AsAtDate As DateTime
    Private FromDate As DateTime
    Private IsBBF As String
    Private IsGAT As String
    Private dsPayment As DataSet

    Private g_StructureFilename As String = ""
    Private g_sReportFilename As String = ""
    Private g_bIsShared As Boolean = False
    Private g_sPVMailRunningDate As String = ""

    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
    Private g_sDocNum As String = ""
    Private g_sDocEntry As String = ""
    Private g_sSeries As String = ""
    Private bolShowDetails As Boolean = False
    Private sShowTaxDate As String = String.Empty

#End Region

#Region "Intialize Application"
    Public Sub New()
        Try
            'Setup_Notes()
        Catch ex As Exception
            MsgBox("[NCM_PVEmail_Param].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        Dim oPictureBox As SAPbouiCOM.PictureBox
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmPV_Email_Param.srf") Then

            oForm = SBO_Application.Forms.Item("ncmPV_Email_Param")
            oPictureBox = oForm.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"

            oForm.Items.Item("lbStatus").FontSize = 10
            'oForm.Items.Item("lbStyleOpt").TextStyle = 4

            AddDataSource()
            SetDatasource()
            SetupChooseFromList()

            oForm.Items.Item("txtBPFr").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Visible = True
        Else
            Try
                oForm = SBO_Application.Forms.Item("ncmPV_Email_Param")
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
            .Add("txtDateFr", SAPbouiCOM.BoDataType.dt_DATE)
            .Add("txtDateTo", SAPbouiCOM.BoDataType.dt_DATE)
            .Add("txtBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtBPGFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtBPGTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtDocFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("txtDocTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
        End With
    End Sub
    Private Sub SetDatasource()
        
        oEdit = oForm.Items.Item("txtDateFr").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtDateFr")
        'oEdit.Value = Now.ToString("yyyyMM01")
        oEdit = oForm.Items.Item("txtDateTo").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtDateTo")

        oEdit = oForm.Items.Item("txtBPFr").Specific
        oEdit.DataBind.SetBound(True, "", "txtBPFr")
        oEdit = oForm.Items.Item("txtBPTo").Specific
        oEdit.DataBind.SetBound(True, "", "txtBPTo")
        oEdit = oForm.Items.Item("txtBPGFr").Specific
        oEdit.DataBind.SetBound(True, "", "txtBPGFr")
        oEdit = oForm.Items.Item("txtBPGTo").Specific
        oEdit.DataBind.SetBound(True, "", "txtBPGTo")
        oEdit = oForm.Items.Item("txtDocFr").Specific
        oEdit.DataBind.SetBound(True, "", "txtDocFr")
        oEdit = oForm.Items.Item("txtDocTo").Specific
        oEdit.DataBind.SetBound(True, "", "txtDocTo")

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
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.PV) & "'"

            g_sReportFilename = GetSharedFilePath(ReportName.PV)
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
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[PaymentVoucher_Email].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub LoadViewer()
        oForm.Items.Item("btnExecute").Enabled = False
        Try
            Dim frm As New Hydac_FormViewer
            Dim sAttachPath As String = ""
            Dim bIsContinue As Boolean = False
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\{1}_PV_{2}.pdf"
            Dim sCurrDate As String = ""
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")

            Try

                g_sPVMailRunningDate = ""
                Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') , IFNULL(""AttachPath"",'')  FROM " & oCompany.CompanyDB & ".""OADP"" ")
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    g_sPVMailRunningDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
                    sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
                    sAttachPath = Convert.ToString(oRec.Fields.Item(2).Value).Trim
                End If

                oRec.DoQuery("SELECT ""U_INVDETAIL"", ""U_TAXDATE"" FROM """ & oCompany.CompanyDB & """.""@NCM_NEW_SETTING"" ")
                bolShowDetails = IIf(oRec.Fields.Item(0).Value = "Y", True, False)
                sShowTaxDate = oRec.Fields.Item(1).Value

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

                'g_sPVMailRunningDate = "2015-02-24 16:52:20.4710"
                ' ===============================================================================
                ' get the folder of PV of the current DB Name
                ' set to local
                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\PVMail\" & oCompany.CompanyDB
                'Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                'If Not di.Exists Then
                '    di.Create()
                'End If
                'sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                'sFinalFileName = di.FullName & "\PVMail_" & sCurrDate & "_" & sCurrTime & ".pdf"
                ' ===============================================================================

                If Save_Settings() Then
                    sQuery = "Select T0.""DocEntry"", T0.""DocNum"", T0.""Series"", T0.""CardCode"", T1.""CardName"", T0.""DocCurr"", T0.""DocTotal"", IFNULL(T1.""U_PV_MailTo"",'') as ""EmailTo""  "
                    sQuery &= " FROM " & oCompany.CompanyDB & ".OVPM T0 "
                    sQuery &= " INNER JOIN " & oCompany.CompanyDB & ".OCRD T1 ON T0.""CardCode"" = T1.""CardCode"" "
                    sQuery &= " INNER JOIN " & oCompany.CompanyDB & ".OCRG T2 ON T1.""GroupCode"" = T2.""GroupCode"" "
                    sQuery &= " WHERE T0.""DocType"" = 'S' AND T0.""Canceled"" = 'N' "

                    oEdit = oForm.Items.Item("txtBPFr").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T1.""CardCode"" >= '" & oEdit.Value.ToString() & "' "
                    End If

                    oEdit = oForm.Items.Item("txtBPTo").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T1.""CardCode"" <= '" & oEdit.Value.ToString() & "' "
                    End If

                    oEdit = oForm.Items.Item("txtBPGFr").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T2.""GroupCode"" >= '" & oEdit.Value.ToString() & "' "
                    End If

                    oEdit = oForm.Items.Item("txtBPGTo").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T2.""GroupCode"" <= '" & oEdit.Value.ToString() & "' "
                    End If

                    oEdit = oForm.Items.Item("txtDocFr").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T0.""DocNum"" >= '" & oEdit.Value.ToString() & "' "
                    End If

                    oEdit = oForm.Items.Item("txtDocTo").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T0.""DocNum"" <= '" & oEdit.Value.ToString() & "' "
                    End If

                    oEdit = oForm.Items.Item("txtDateFr").Specific
                    FromDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("txtDateFr").ValueEx, "yyyyMMdd", Nothing)
                    If oEdit.Value.ToString().Trim() <> "" Then
                        sQuery &= " AND T0.""DocDate"" >= '" & FromDate.ToString("yyyyMMdd") & "' "
                    End If

                    oEdit = oForm.Items.Item("txtDateTo").Specific
                    If oEdit.Value.ToString().Trim() <> "" Then
                        FromDate = DateTime.ParseExact(oForm.DataSources.UserDataSources.Item("txtDateTo").ValueEx, "yyyyMMdd", Nothing)
                        sQuery &= " AND T0.""DocDate"" <= '" & FromDate.ToString("yyyyMMdd") & "' "
                    End If

                    sQuery &= " Order by T0.""CardCode"", ""DocNum"" "

                    oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)

                    If oRec.RecordCount > 0 Then
                        ShowStatus("Preparing the email, please wait...")
                        Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                        If Not di.Exists Then
                            di.Create()
                        End If

                        Dim ds As New dsPVEmail()

                        While Not oRec.EoF
                            g_sDocEntry = oRec.Fields.Item("DocEntry").Value.ToString()
                            g_sDocNum = oRec.Fields.Item("DocNum").Value.ToString()
                            g_sSeries = oRec.Fields.Item("Series").Value.ToString()
                            If PrepareDataset() Then
                                With frm
                                    .ReportName = ReportName.PV
                                    .ExportPath = sFinalFileName
                                    .Dataset = dsPayment
                                    .DocNum = g_sDocNum
                                    .DocEntry = g_sDocEntry
                                    .Series = g_sSeries
                                    .DBUsernameViewer = DBUsername
                                    .DBPasswordViewer = DBPassword
                                    .ShowDetails = bolShowDetails
                                    .ShowTaxDate = sShowTaxDate
                                    .IsShared = g_bIsShared
                                    .ReportNamePV = g_sReportFilename
                                    .DatabaseServer = oCompany.Server
                                    .DatabaseName = oCompany.CompanyDB
                                    .IsExport = True
                                    .CrystalReportExportType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                                    .CrystalReportExportPath = String.Format(sPathFormat, di.FullName, oRec.Fields.Item("CardCode").Value.ToString(), System.DateTime.Now.Date.ToString("ddMMyyyy"))
                                    frm.OPEN_HANADS_PV_EMAIL()

                                End With

                                Dim dr As dsPVEmail.PreviewDTRow
                                dr = ds.PreviewDT.NewPreviewDTRow()
                                dr.Attachment = String.Format(sPathFormat, di.FullName, oRec.Fields.Item("CardCode").Value, System.DateTime.Now.Date.ToString("ddMMyyyy"))
                                dr.Balance = oRec.Fields.Item("DocTotal").Value
                                dr.CardCode = oRec.Fields.Item("CardCode").Value
                                dr.CardName = oRec.Fields.Item("CardName").Value
                                dr.Currency = oRec.Fields.Item("DocCurr").Value
                                dr.EmailTo = oRec.Fields.Item("EmailTo").Value
                                dr.DocEntry = oRec.Fields.Item("DocEntry").Value
                                dr.DocNum = oRec.Fields.Item("DocNum").Value
                                dr.IsEmail = IIf(dr.Balance > 0, 1, 0)

                                dr.Table.Rows.Add(dr)

                            End If

                            oRec.MoveNext()
                        End While

                        ShowStatus("Showing email list, please wait...")

                        If ds.Tables(0).Rows.Count > 0 Then
                            SubMain.oFrmPVSendEmail.ReportName = ReportName.PV
                            SubMain.oFrmPVSendEmail.StatementAsAtDate = AsAtDate
                            SubMain.oFrmPVSendEmail.StatementDataTable = ds.PreviewDT
                            SubMain.oFrmPVSendEmail.LoadForm()
                            Hydac_FormViewer.Close()
                        End If

                    Else
                        ShowStatus("There no data")
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item("btnExecute").Enabled = True
            End Try
            
        Catch ex As Exception
            SBO_Application.MessageBox("[PVMail].[LoadViewer]:" & ex.Message)
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
            oCFLCreation.ObjectType = 46
            oCFLCreation.UniqueID = "CFL_DocFrom"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtDocFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_DocFrom"
            oEditLn.ChooseFromListAlias = "DocNum"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 46
            oCFLCreation.UniqueID = "CFL_DocTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtDocTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_DocTo"
            oEditLn.ChooseFromListAlias = "DocNum"
            ' ----------------------------------------
        Catch ex As Exception
            Throw New Exception("[NCM_PV_Email_Param].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function PrepareDataset() As Boolean
        Try
            If g_StructureFilename.Length <= 0 Then
                dsPayment = New DS_PAYMENT
            Else
                dsPayment = New DataSet
                dsPayment.ReadXml(g_StructureFilename)
            End If

            Dim ProviderName As String = "System.Data.Odbc"
            Dim sQuery As String = ""
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            Dim dtOADM As System.Data.DataTable
            Dim dtADM1 As System.Data.DataTable
            Dim dtIMAGE As System.Data.DataTable
            Dim dtNNM1 As System.Data.DataTable
            Dim dtOACT As System.Data.DataTable

            Dim dtOVPM As System.Data.DataTable
            Dim dtVPM1 As System.Data.DataTable
            Dim dtVPM2 As System.Data.DataTable
            Dim dtVPM3 As System.Data.DataTable
            Dim dtVPM4 As System.Data.DataTable

            Dim dtNNM1_1 As System.Data.DataTable
            Dim dtNNM1_2 As System.Data.DataTable
            Dim dtNNM1_3 As System.Data.DataTable
            Dim dtNNM1_4 As System.Data.DataTable
            Dim dtNNM1_5 As System.Data.DataTable
            Dim dtNNM1_6 As System.Data.DataTable
            Dim dtNNM1_7 As System.Data.DataTable

            Dim dtOJDT As System.Data.DataTable
            Dim dtOINV As System.Data.DataTable
            Dim dtORIN As System.Data.DataTable
            Dim dtOPCH As System.Data.DataTable
            Dim dtORPC As System.Data.DataTable
            Dim dtODPO As System.Data.DataTable
            Dim dtODPI As System.Data.DataTable

            Dim dtINV1 As System.Data.DataTable
            Dim dtRIN1 As System.Data.DataTable
            Dim dtPCH1 As System.Data.DataTable
            Dim dtRPC1 As System.Data.DataTable
            Dim dtDPO1 As System.Data.DataTable
            Dim dtDPI1 As System.Data.DataTable

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            '------INV HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OINV"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '13') "
            dtOINV = dsPayment.Tables("OINV")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOINV)

            '------INV LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""INV1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '13') "
            dtINV1 = dsPayment.Tables("INV1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtINV1)

            '------RIN HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ORIN"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '14') "
            dtORIN = dsPayment.Tables("ORIN")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtORIN)

            '------RIN LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""RIN1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '14') "
            dtRIN1 = dsPayment.Tables("RIN1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRIN1)

            '------PCH HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OPCH"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '18') "
            dtOPCH = dsPayment.Tables("OPCH")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOPCH)

            '------PCH LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""PCH1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '18') "
            dtPCH1 = dsPayment.Tables("PCH1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPCH1)

            '------RPC HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ORPC"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '19') "
            dtORPC = dsPayment.Tables("ORPC")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtORPC)

            '------RPC LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""RPC1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '19') "
            dtRPC1 = dsPayment.Tables("RPC1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRPC1)

            '------DPI HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ODPI"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '203') "
            dtODPI = dsPayment.Tables("ODPI")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtODPI)

            '------DPI LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""DPI1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '203') "
            dtDPI1 = dsPayment.Tables("DPI1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtDPI1)

            '------DPO HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ODPO"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '204') "
            dtODPO = dsPayment.Tables("ODPO")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtODPO)

            '------DPO LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""DPO1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '204') "
            dtDPO1 = dsPayment.Tables("DPO1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtDPO1)

            '------JE--------------------------------------------------
            sQuery = " SELECT ""LogInstanc"",""DocSeries"",""DocType"",""DueDate"",""Number"",""Memo"",""ObjType"",""Ref1"",""Ref2"",""Ref3"",""RefDate"",""Series"",""SeriesStr"",""TaxDate"",""TransCode"",  CASE WHEN IFNULL(""TransCurr"",'') = '' THEN (SELECT ""MainCurncy"" FROM """ & oCompany.CompanyDB & """.""OADM"") ELSE ""TransCurr"" END ""TransCurr"",""TransId"",""TransType"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OJDT"" "
            sQuery &= " WHERE ""TransType"" = '30' AND ""TransId"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '30') "
            sQuery &= " UNION ALL "
            sQuery &= " SELECT ""LogInstanc"",""DocSeries"",""DocType"",""DueDate"",""Number"",""Memo"",""ObjType"",""Ref1"",""Ref2"",""Ref3"",""RefDate"",""Series"",""SeriesStr"",""TaxDate"",""TransCode"", CASE WHEN IFNULL(""TransCurr"",'') = '' THEN (SELECT ""MainCurncy"" FROM """ & oCompany.CompanyDB & """.""OADM"") ELSE ""TransCurr"" END ""TransCurr"",""TransId"",""TransType"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OJDT"" "
            sQuery &= " WHERE ""TransType"" = '24' AND ""TransId"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '24') "
            sQuery &= " UNION ALL "
            sQuery &= " SELECT ""LogInstanc"",""DocSeries"",""DocType"",""DueDate"",""Number"",""Memo"",""ObjType"",""Ref1"",""Ref2"",""Ref3"",""RefDate"",""Series"",""SeriesStr"",""TaxDate"",""TransCode"", CASE WHEN IFNULL(""TransCurr"",'') = '' THEN (SELECT ""MainCurncy"" FROM """ & oCompany.CompanyDB & """.""OADM"") ELSE ""TransCurr"" END ""TransCurr"",""TransId"",""TransType"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OJDT"" "
            sQuery &= " WHERE ""TransType"" = '46' AND ""TransId"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '46') "

            dtOJDT = dsPayment.Tables("OJDT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOJDT)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""OVPM"" WHERE ""DocNum"" = '" & g_sDocNum & "' AND ""Series"" = '" & g_sSeries & "' AND ""DocEntry"" = '" & g_sDocEntry & "' "
            dtOVPM = dsPayment.Tables("OVPM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOVPM)

            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM1"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM1 = dsPayment.Tables("VPM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM1)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM2"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM2 = dsPayment.Tables("VPM2")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM2)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM3"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM3 = dsPayment.Tables("VPM3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM3)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM4"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM4 = dsPayment.Tables("VPM4")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM4)

            '--------------------------------------------------------
            sQuery = " SELECT '1' ""FLAG"", '1' ""SRNO"" FROM DUMMY "
            dtIMAGE = dsPayment.Tables("@NCM_IMAGE")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtIMAGE)

            '--------------------------------------------------------
            sQuery = " SELECT ""AcctCode"",""AcctName"",""LogInstanc"",""FormatCode"" FROM """ & oCompany.CompanyDB & """.""OACT"" "
            dtOACT = dsPayment.Tables("OACT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOACT)

            '--------------------------------------------------------
            sQuery = "  SELECT ""ObjectCode"", ""Series"", ""SeriesName"", IFNULL(""BeginStr"",'') AS ""BeginStr"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '46' "
            sQuery &= " UNION ALL "
            sQuery &= " SELECT '46' ""ObjectCode"", '-1' ""Series"", 'Manual' ""SeriesName"", '' ""BeginStr""  "
            sQuery &= " FROM ""DUMMY"" "
            dtNNM1 = dsPayment.Tables("NNM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1)

            '--------------------------------------------------------
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"" FROM """ & oCompany.CompanyDB & """.""ADM1""  "
            dtADM1 = dsPayment.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)

            '--------------------------------------------------------
            sQuery = "SELECT ""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            dtOADM = dsPayment.Tables("OADM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)

            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '18' "
            dtNNM1_1 = dsPayment.Tables("NCM_NNM1_1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_1)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '19' "
            dtNNM1_2 = dsPayment.Tables("NCM_NNM1_2")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_2)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" IN ('24','46','30') "
            dtNNM1_3 = dsPayment.Tables("NCM_NNM1_3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_3)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '204' "
            dtNNM1_4 = dsPayment.Tables("NCM_NNM1_4")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_4)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '13' "
            dtNNM1_5 = dsPayment.Tables("NCM_NNM1_5")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_5)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '203' "
            dtNNM1_6 = dsPayment.Tables("NCM_NNM1_6")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_6)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '14' "
            dtNNM1_7 = dsPayment.Tables("NCM_NNM1_7")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_7)
            '--------------------------------------------------------
            dbConn.Close()

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PrepareDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

#Region "Logic Function"

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

            'sQuery = "UPDATE """ & oCompany.CompanyDB & """.""@NCM_SOC2"" SET ""NOTES"" ='" & Notes & "' WHERE ""ID"" = '1'"
            'oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRec.DoQuery(sQuery)

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[PV Mail].[Save_Settings]:" & ex.Message)
            Return False
        End Try
    End Function
    
    Private Function ValidateParameter() As Boolean
        Try
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

            sStart = oForm.DataSources.UserDataSources.Item("txtDocFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtDocTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Doc. No from is greater than Doc. No to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtDocFr"
                    Return False
                End If
            End If

            If (oForm.DataSources.UserDataSources.Item("txtDateFr").ValueEx.Length = 0) Then
                SBO_Application.StatusBar.SetText("Please enter from date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "txtDateFr"
                Return False
            End If


            SBO_Application.StatusBar.SetText(String.Empty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[PV Mail].[ValidateParameter] - " & ex.Message, 1, "OK", String.Empty, String.Empty)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Friend Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                If pVal.FormUID = "ncmPV_Email_Param" Then
                    Select Case pVal.ItemUID
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
                            Case "CFL_BGFrom"
                                sTemp = oDataTable.GetValue("GroupCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtBPGFr").ValueEx = sTemp
                                Exit Select
                            Case "CFL_BGTo"
                                sTemp = oDataTable.GetValue("GroupCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtBPGTo").ValueEx = sTemp
                                Exit Select
                            Case "CFL_DocFrom"
                                sTemp = oDataTable.GetValue("DocNum", 0)
                                oForm.DataSources.UserDataSources.Item("txtDocFr").ValueEx = sTemp
                                Exit Select
                            Case "CFL_DocTo"
                                sTemp = oDataTable.GetValue("DocNum", 0)
                                oForm.DataSources.UserDataSources.Item("txtDocTo").ValueEx = sTemp
                                Exit Select
                            Case Else
                                Exit Select
                        End Select
                        Return True
                    End If
                End If

                If pVal.FormUID = "ncmPV_Email_Param" Then
                    Select Case pVal.ItemUID
                        Case "btnExecute"
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                    Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                    myThread.SetApartmentState(Threading.ApartmentState.STA)
                                    myThread.Start()
                                End If
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[NCM_PV_Email_Param].[ItemEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region
End Class
