Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class PaymentDraft
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix    'JN added
    Private g_sDocEntry As String = ""
    Private g_sDocNum As String = ""
    Private g_sSeries As Integer = 0   'JN change
    Private dsPAYMENT As DataSet
    Private g_bIsShared As Boolean = False

    Private g_StructureFilename As String = ""
    Private strDocType As String = ""
    Private ObjType As String = ""
    Private DocTime As String = ""
    Private g_sReportType As String = ""
    Private g_sReportFilename As String = ""
    Private bolShowDetails As Boolean = False
    Private sShowTaxDate As String = String.Empty

    Friend Function IsTriggerPV() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            sQuery = "  SELECT TOP 1 IFNULL(""INCLUDED"",'N') FROM """ & oCompany.CompanyDB & """.""@NCM_RPT_CONFIG"" "
            sQuery &= " WHERE ""RPTCODE"" = '" & GetReportCode(ReportName.PVDraft) & "'"
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                If oRec.Fields.Item(0).Value = "Y" Then
                    oRec = Nothing
                    Return True
                End If
            End If
            oRec = Nothing
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PaymentDraft].[IsTriggerPV] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            g_sReportFilename = ""
            g_StructureFilename = ""

            sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.PVDraft) & "'"

            g_sReportFilename = GetSharedFilePath(ReportName.PVDraft)
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
            SBO_Application.StatusBar.SetText("[PaymentDraft].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function PrepareDataset() As Boolean
        Try
            If g_StructureFilename.Length <= 0 Then
                dsPAYMENT = New DS_PAYMENT
            Else
                dsPAYMENT = New DataSet
                dsPAYMENT.ReadXml(g_StructureFilename)
            End If

            Dim ProviderName As String = "System.Data.Odbc"
            Dim sQuery As String = ""
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            Dim dtNNM1 As System.Data.DataTable
            Dim dtOADM As System.Data.DataTable
            Dim dtADM1 As System.Data.DataTable
            Dim dtOACT As System.Data.DataTable
            Dim dtVIEW As System.Data.DataTable
            Dim dtOPDF As System.Data.DataTable
            Dim dtPDF1 As System.Data.DataTable
            Dim dtPDF3 As System.Data.DataTable
            Dim dtPDF4 As System.Data.DataTable

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            '--------------------------------------------------------
            sQuery = " SELECT ""ObjectCode"",""Series"",""SeriesName"",IFNULL(""BeginStr"",'') AS ""BeginStr"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '46' "
            dtNNM1 = dsPAYMENT.Tables("NNM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1)
            '--------------------------------------------------------
            sQuery = "  Select T1.*, T2.""CardName"", T2.""BankCode"", T3.""BankName"", T4.""INTERNAL_K"", T4.""U_NAME"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OPDF"" T1 "
            sQuery &= " LEFT OUTER JOIN  """ & oCompany.CompanyDB & """.""OCRD"" T2 On T1.""CardCode"" = T2.""CardCode"" "
            sQuery &= " LEFT OUTER JOIN  """ & oCompany.CompanyDB & """.""ODSC"" T3 On T2.""BankCode"" = T3.""BankCode"" "
            sQuery &= " LEFT OUTER JOIN  """ & oCompany.CompanyDB & """.""OUSR"" T4 On T1.""UserSign"" = T4.""INTERNAL_K"" "
            sQuery &= " WHERE T1.""DocNum"" = '" & g_sDocNum & "' AND T1.""DocEntry"" = '" & g_sDocEntry & "' AND T1.""Series"" = '" & g_sSeries & "'"

            dtOPDF = dsPAYMENT.Tables("OPDF")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOPDF)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""PDF1"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtPDF1 = dsPAYMENT.Tables("PDF1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPDF1)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""PDF3"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtPDF3 = dsPAYMENT.Tables("PDF3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPDF3)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""PDF4"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtPDF4 = dsPAYMENT.Tables("PDF4")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPDF4)
            '--------------------------------------------------------
            sQuery = " SELECT ""AcctCode"",""AcctName"",""LogInstanc"",""FormatCode"" FROM """ & oCompany.CompanyDB & """.""OACT"" "
            dtOACT = dsPAYMENT.Tables("OACT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOACT)
            '--------------------------------------------------------
            sQuery = "SELECT ""FaxF"",""Phone1F"",""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"",""DdctOffice"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            dtOADM = dsPAYMENT.Tables("OADM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)
            '--------------------------------------------------------
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"", ""StreetF"", ""BlockF"", ""ZipCodeF"", ""BuildingF"" FROM """ & oCompany.CompanyDB & """.""ADM1""  "
            dtADM1 = dsPAYMENT.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)
            '--------------------------------------------------------

            sQuery = "SELECT * FROM """ & oCompany.CompanyDB & """.""NCM_VIEW_DRAFTPV_INVOICE"" WHERE ""PaymentDocEntry"" = '" & g_sDocEntry & "' AND ""PaymentObjType"" = '46' "
            dtVIEW = dsPAYMENT.Tables("NCM_VIEW_DRAFTPV_INVOICE")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVIEW)
            '--------------------------------------------------------

            dbConn.Close()

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PrepareDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Friend Sub LoadViewer()
        Try
            Dim frm As Hydac_FormViewer = New Hydac_FormViewer
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\DOP_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oRec.DoQuery("SELECT  TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                sCurrDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
            End If

            ' ===============================================================================
            ' get the folder of the current DB Name
            ' set to local
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\DOP\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\DOP_" & sCurrDate & "_" & sCurrTime & ".pdf"
            ' ===============================================================================

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

            If PrepareDataset() Then
                With frm
                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            .ClientType = "D"
                        Case SAPbouiCOM.BoClientType.ct_Browser
                            .ClientType = "S"
                    End Select

                    .ExportPath = sFinalFileName
                    .Dataset = dsPAYMENT
                    .Text = "Draft Payment Voucher - " & g_sDocNum
                    .ReportName = ReportName.PVDraft
                    .DocNum = g_sDocNum
                    .Series = g_sSeries
                    .DocEntry = g_sDocEntry
                    .DBUsernameViewer = DBUsername
                    .DBPasswordViewer = DBPassword
                    .ShowDetails = bolShowDetails
                    .ShowTaxDate = sShowTaxDate
                    .IsShared = g_bIsShared
                    .ReportNamePV = g_sReportFilename

                End With

                Select Case SBO_Application.ClientType
                    Case SAPbouiCOM.BoClientType.ct_Desktop
                        frm.ShowDialog()

                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.OPEN_HANADS_DRAFTPV()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select
            End If
           
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PaymentDraft].[LoadViewer]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error) 'JN changed
        End Try
    End Sub
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.Before_Action = False Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                        oForm = SBO_Application.Forms.GetForm(pval.FormType, pval.FormTypeCount)
                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        oForm = SBO_Application.Forms.GetForm(pval.FormType, pval.FormTypeCount)
                End Select
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PaymentDraft].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)   'JN changed
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
    Public Function MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Dim iVal As Integer
        Try
            If pval.BeforeAction = True Then
                Select Case pval.MenuUID
                    Case MenuID.PrintPreview
                        If (IsTriggerPV()) Then
                            BubbleEvent = False
                            oMatrix = oForm.Items.Item("5").Specific  'JN added

                            Dim oDB As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("OPDF")
                            iVal = oMatrix.GetNextSelectedRow(0)   'JN added

                            oMatrix.GetLineData(iVal)
                            g_sDocNum = oDB.GetValue("DocNum", oDB.Offset).ToString.Trim
                            g_sSeries = oDB.GetValue("Series", oDB.Offset).ToString.Trim
                            ObjType = oDB.GetValue("ObjType", oDB.Offset).ToString.Trim
                            DocTime = oDB.GetValue("DocTime", oDB.Offset).ToString.Trim
                            g_sDocEntry = oDB.GetValue("DocEntry", oDB.Offset).ToString.Trim

                            Dim oRecordset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            Dim sCheck As String = "SELECT ""DocEntry"", ""DocType"", ""CashSum"", ""TrsfrSum"" FROM """ & oCompany.CompanyDB & """.""OPDF"" WHERE ""DocEntry"" = '" & g_sDocEntry & "' AND ""Series"" = '" & g_sSeries & "' AND ""DocNum"" = '" & g_sDocNum & "' AND ""ObjType"" = '" & ObjType & "' AND ""DocTime"" ='" & DocTime & "'"

                            oRecordset.DoQuery(sCheck)
                            If oRecordset.RecordCount > 0 Then
                                ' --------------------------------------------------------------------------------------------
                                Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRec.DoQuery("SELECT ""U_INVDETAIL"", ""U_TAXDATE"" FROM """ & oCompany.CompanyDB & """.""@NCM_NEW_SETTING"" ")
                                bolShowDetails = IIf(oRec.Fields.Item(0).Value = "Y", True, False)
                                sShowTaxDate = oRec.Fields.Item(1).Value

                                If bolShowDetails = True Then
                                    If oRecordset.Fields.Item("DocType").Value = "S" Then
                                        If oRecordset.Fields.Item("CashSum").Value <> 0 Or oRecordset.Fields.Item("TrsfrSum").Value <> 0 Then
                                            If SBO_Application.MessageBox("Do you want to print Invoice Details?", 1, "Yes", "No") = 1 Then
                                                bolShowDetails = True
                                            Else
                                                bolShowDetails = False
                                            End If
                                            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                                            myThread.Start()
                                        Else
                                            bolShowDetails = False
                                            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                                            myThread.Start()
                                        End If
                                    Else
                                        bolShowDetails = False
                                        Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                        myThread.SetApartmentState(Threading.ApartmentState.STA)
                                        myThread.Start()
                                    End If
                                Else
                                    bolShowDetails = False
                                    Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                    myThread.SetApartmentState(Threading.ApartmentState.STA)
                                    myThread.Start()
                                End If
                                ' --------------------------------------------------------------------------------------------
                            End If
                            BubbleEvent = False
                        Else
                            BubbleEvent = True
                        End If
                End Select
            Else
                ' do nothing
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[PaymentDraft].[MenuEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
End Class