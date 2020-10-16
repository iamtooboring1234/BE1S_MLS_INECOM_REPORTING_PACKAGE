Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class IncomingPayment
    Private oForm As SAPbouiCOM.Form
    Private oCombo As SAPbouiCOM.ComboBox
    Private oEdit As SAPbouiCOM.EditText
    Private g_sDocEntry As String = ""
    Private g_sDocNum As String = ""
    Private g_sSeries As String = ""
    Private dsPAYMENT As DataSet

    Private g_bIsShared As Boolean = False
    Private g_StructureFilename As String = ""
    Private g_sReportFilename As String = ""
    Private bolShowDetails As Boolean = False
    Private sShowTaxDate As String = String.Empty

    Friend Function IsTriggerOfficialReceipt() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            sQuery = "  SELECT IFNULL(""INCLUDED"",'N') FROM ""@NCM_RPT_CONFIG"" "
            sQuery &= " WHERE ""RPTCODE"" = '" & GetReportCode(ReportName.IRA) & "'"
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                If oRec.Fields.Item(0).Value = "Y" Then
                    Return True
                End If
            Else
                sQuery = "  SELECT IFNULL(""INCLUDED"",'N') FROM ""@NCM_RPT_CONFIG"" "
                sQuery &= " WHERE ""RPTCODE"" = '" & GetReportCode(ReportName.OfficialReceipt) & "'"
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    If oRec.Fields.Item(0).Value = "Y" Then
                        Return True
                    End If
                End If
            End If

            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[IsTriggerIRA] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.IRA) & "'"

            g_sReportFilename = GetSharedFilePath(ReportName.IRA)
            If g_sReportFilename = "" Then
                sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
                sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.OfficialReceipt) & "'"

                g_sReportFilename = GetSharedFilePath(ReportName.OfficialReceipt)
            End If

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
            SBO_Application.StatusBar.SetText("[Receipt].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Friend Sub LoadViewer()
        Try
            Dim frm As Hydac_FormViewer = New Hydac_FormViewer
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\IP_{1}.pdf"
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
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\IP\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\IP_" & sCurrDate & sCurrTime & "_" & g_sDocNum & ".pdf"
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
                    .Text = "Official Receipt - " & g_sDocNum
                    .ReportName = ReportName.IRA
                    .DocNum = g_sDocNum
                    .DocEntry = g_sDocEntry
                    .Series = g_sSeries
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
                        frm.OPEN_HANADS_OFFICIALRECEIPT()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[OfficialReceipt].[LoadViewer]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
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

            Dim dtOADM As System.Data.DataTable
            Dim dtADM1 As System.Data.DataTable
            Dim dtNNM1 As System.Data.DataTable
            Dim dtOACT As System.Data.DataTable
            Dim dtOUSR As System.Data.DataTable

            Dim dtORCT As System.Data.DataTable
            Dim dtRCT1 As System.Data.DataTable
            Dim dtRCT2 As System.Data.DataTable
            Dim dtRCT3 As System.Data.DataTable
            Dim dtRCT4 As System.Data.DataTable

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
            sQuery = "  SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OINV"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '13') "
            dtOINV = dsPAYMENT.Tables("OINV")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOINV)

            '------INV LINE--------------------------------------------------
            sQuery = "  SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""INV1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '13') "
            dtINV1 = dsPAYMENT.Tables("INV1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtINV1)

            '------RIN HEADER--------------------------------------------------
            sQuery = "  SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ORIN"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '14') "
            dtORIN = dsPAYMENT.Tables("ORIN")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtORIN)

            '------RIN LINE--------------------------------------------------
            sQuery = "  SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""RIN1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '14') "
            dtRIN1 = dsPAYMENT.Tables("RIN1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRIN1)

            '------PCH HEADER--------------------------------------------------
            sQuery = "  SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OPCH"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '18') "
            dtOPCH = dsPAYMENT.Tables("OPCH")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOPCH)

            '------PCH LINE--------------------------------------------------
            sQuery = "  SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""PCH1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '18') "
            dtPCH1 = dsPAYMENT.Tables("PCH1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPCH1)

            '------RPC HEADER--------------------------------------------------
            sQuery = "  SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ORPC"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '19') "
            dtORPC = dsPAYMENT.Tables("ORPC")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtORPC)

            '------RPC LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""RPC1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '19') "
            dtRPC1 = dsPAYMENT.Tables("RPC1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRPC1)

            '------DPI HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ODPI"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '203') "
            dtODPI = dsPAYMENT.Tables("ODPI")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtODPI)

            '------DPI LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""DPI1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '203') "
            dtDPI1 = dsPAYMENT.Tables("DPI1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtDPI1)

            '------DPO HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ODPO"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '204') "
            dtODPO = dsPAYMENT.Tables("ODPO")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtODPO)

            '------DPO LINE--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""LineNum"",""VisOrder"",""ItemCode"",""Dscription"",""Quantity"",""Price"",""TotalFrgn"",""Rate"",""LineTotal"" "
            sQuery &= "  FROM """ & oCompany.CompanyDB & """.""DPO1"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '204') "
            dtDPO1 = dsPAYMENT.Tables("DPO1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtDPO1)

            '------JE--------------------------------------------------
            sQuery = " SELECT ""LogInstanc"",""DocSeries"",""DocType"",""DueDate"",""Number"",""Memo"",""ObjType"",""Ref1"",""Ref2"",""Ref3"",""RefDate"",""Series"",""SeriesStr"",""TaxDate"",""TransCode"", CASE WHEN IFNULL(""TransCurr"",'') = '' THEN (SELECT ""MainCurncy"" FROM """ & oCompany.CompanyDB & """.""OADM"") ELSE ""TransCurr"" END ""TransCurr"" ,""TransId"",""TransType"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OJDT"" "
            sQuery &= " WHERE ""TransType"" = '30' AND ""TransId"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '30') "
            sQuery &= " UNION ALL "
            sQuery &= " SELECT ""LogInstanc"",""DocSeries"",""DocType"",""DueDate"",""Number"",""Memo"",""ObjType"",""Ref1"",""Ref2"",""Ref3"",""RefDate"",""Series"",""SeriesStr"",""TaxDate"",""TransCode"", CASE WHEN IFNULL(""TransCurr"",'') = '' THEN (SELECT ""MainCurncy"" FROM """ & oCompany.CompanyDB & """.""OADM"") ELSE ""TransCurr"" END ""TransCurr"" ,""TransId"",""TransType"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OJDT"" "
            sQuery &= " WHERE ""TransType"" = '24' AND ""TransId"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '24') "
            sQuery &= " UNION ALL "
            sQuery &= " SELECT ""LogInstanc"",""DocSeries"",""DocType"",""DueDate"",""Number"",""Memo"",""ObjType"",""Ref1"",""Ref2"",""Ref3"",""RefDate"",""Series"",""SeriesStr"",""TaxDate"",""TransCode"", CASE WHEN IFNULL(""TransCurr"",'') = '' THEN (SELECT ""MainCurncy"" FROM """ & oCompany.CompanyDB & """.""OADM"") ELSE ""TransCurr"" END ""TransCurr"" ,""TransId"",""TransType"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OJDT"" "
            sQuery &= " WHERE ""TransType"" = '46' AND ""TransId"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""RCT2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '46') "

            dtOJDT = dsPAYMENT.Tables("OJDT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOJDT)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""ORCT"" WHERE ""DocNum"" = '" & g_sDocNum & "' AND ""Series"" = '" & g_sSeries & "' AND ""DocEntry"" = '" & g_sDocEntry & "' "
            dtORCT = dsPAYMENT.Tables("ORCT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtORCT)

            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""RCT1"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtRCT1 = dsPAYMENT.Tables("RCT1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRCT1)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""RCT2"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtRCT2 = dsPAYMENT.Tables("RCT2")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRCT2)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""RCT3"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtRCT3 = dsPAYMENT.Tables("RCT3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRCT3)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""RCT4"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtRCT4 = dsPAYMENT.Tables("RCT4")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRCT4)

            '--------------------------------------------------------
            sQuery = " SELECT ""AcctCode"",""AcctName"",""LogInstanc"",""FormatCode"",""Segment_0"",""Segment_1"",""Segment_2"",""Segment_3"",""Segment_4"",""Segment_5"",""Segment_6"",""Segment_7"",""Segment_8"",""Segment_9"" FROM """ & oCompany.CompanyDB & """.""OACT"" "
            dtOACT = dsPAYMENT.Tables("OACT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOACT)

            '--------------------------------------------------------
            sQuery = "SELECT  ""INTERNAL_K"", ""U_NAME"" FROM """ & oCompany.CompanyDB & """.""OUSR""  "
            dtOUSR = dsPAYMENT.Tables("OUSR")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOUSR)
            '--------------------------------------------------------
            sQuery = "  SELECT ""ObjectCode"",""Series"",""SeriesName"",IFNULL(""BeginStr"",'') AS ""BeginStr"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '24' "
            sQuery &= " UNION ALL "
            sQuery &= " SELECT '24' ""ObjectCode"", '-1' ""Series"", 'Manual' ""SeriesName"", '' ""BeginStr""  "
            sQuery &= " FROM ""DUMMY"" "
            dtNNM1 = dsPAYMENT.Tables("NNM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1)

            '--------------------------------------------------------
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"" FROM """ & oCompany.CompanyDB & """.""ADM1""  "
            dtADM1 = dsPAYMENT.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)

            '--------------------------------------------------------
            sQuery = "SELECT ""FaxF"",""Phone1F"",""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            dtOADM = dsPAYMENT.Tables("OADM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)

            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '18' "
            dtNNM1_1 = dsPAYMENT.Tables("NCM_NNM1_1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_1)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '19' "
            dtNNM1_2 = dsPAYMENT.Tables("NCM_NNM1_2")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_2)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" IN ('24','46','30') "
            dtNNM1_3 = dsPAYMENT.Tables("NCM_NNM1_3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_3)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '204' "
            dtNNM1_4 = dsPAYMENT.Tables("NCM_NNM1_4")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_4)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '13' "
            dtNNM1_5 = dsPAYMENT.Tables("NCM_NNM1_5")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_5)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '203' "
            dtNNM1_6 = dsPAYMENT.Tables("NCM_NNM1_6")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1_6)
            '--------------------------------------------------------
            sQuery = "SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM """ & oCompany.CompanyDB & """.""NNM1"" WHERE ""ObjectCode"" = '14' "
            dtNNM1_7 = dsPAYMENT.Tables("NCM_NNM1_7")
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
            SBO_Application.StatusBar.SetText("[PaymentVoucher].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
    Public Function MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = True Then
                Select Case pval.MenuUID
                    Case MenuID.PrintPreview
                        If (IsTriggerOfficialReceipt()) Then
                            BubbleEvent = False
                            Dim oDB As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("ORCT")
                            g_sDocEntry = ""
                            g_sDocEntry = oDB.GetValue("DocEntry", oDB.Offset).ToString.Trim


                            oCombo = oForm.Items.Item("87").Specific
                            g_sSeries = oCombo.Selected.Value
                            oEdit = oForm.Items.Item("3").Specific
                            g_sDocNum = oEdit.Value
                            Dim oRecordset As SAPbobsCOM.Recordset
                            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordset.DoQuery("SELECT ""DocType"", ""CashSum"", ""TrsfrSum"" FROM """ & oCompany.CompanyDB & """.""ORCT"" WHERE ""Series"" = '" & g_sSeries & "' AND ""DocNum"" ='" & g_sDocNum & "' AND ""DocEntry"" = '" & g_sDocEntry & "'")
                            If oRecordset.RecordCount > 0 Then
                                '' --------------------------------------------------------------------------------------------
                                Dim oRec As SAPbobsCOM.Recordset
                                Dim sRec As String = ""
                                sRec = "  SELECT IFNULL(""U_IRAINVDETAIL"",'N'), IFNULL(""U_IRATAXDATE"",'N') "
                                sRec &= " FROM """ & oCompany.CompanyDB & """.""@NCM_NEW_SETTING"" "

                                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRec.DoQuery(sRec)
                                bolShowDetails = IIf(oRec.Fields.Item(0).Value = "Y", True, False)
                                sShowTaxDate = oRec.Fields.Item(1).Value

                                If bolShowDetails = True Then
                                    If oRecordset.Fields.Item("DocType").Value = "C" Then
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
                Select Case pval.MenuUID
                End Select
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[OfficialReceipt].[MenuEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
End Class