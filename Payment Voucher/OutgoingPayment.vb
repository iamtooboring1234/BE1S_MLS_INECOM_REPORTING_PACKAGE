Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class OutgoingPayment
    Private oForm As SAPbouiCOM.Form
    Private oCombo As SAPbouiCOM.ComboBox
    Private oEdit As SAPbouiCOM.EditText
    Private g_sDocNum As String = ""
    Private g_sDocEntry As String = ""
    Private g_sSeries As String = ""
    Private dsPAYMENT As DataSet

    Private g_StructureFilename As String = ""
    Private g_sReportType As String = ""
    Private g_sReportFilename As String = ""
    Private g_bIsShared As Boolean = False
    Private bolShowDetails As Boolean = False
    Private sShowTaxDate As String = String.Empty

    Private Function IsSharedFileExist() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            g_sReportFilename = ""
            g_StructureFilename = ""

            Select Case g_sReportType
                Case "PV"
                    sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
                    sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.PV) & "'"

                    g_sReportFilename = GetSharedFilePath(ReportName.PV)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            'okay
                        End If
                    End If
                Case "RA"
                    sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
                    sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.RA) & "'"

                    g_sReportFilename = GetSharedFilePath(ReportName.RA)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            'okay
                        End If
                    End If
            End Select

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
            SBO_Application.StatusBar.SetText("[PaymentVoucher].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Friend Function IsTriggerPV() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""
            sQuery = "  SELECT IFNULL(""INCLUDED"",'N') FROM """ & oCompany.CompanyDB & """.""@NCM_RPT_CONFIG"" WHERE ""RPTCODE"" = '" & GetReportCode(ReportName.PV) & "'"
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                If oRec.Fields.Item(0).Value = "Y" Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[IsTriggerPV] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"", ""Comments"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OINV"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '13') "
            dtOINV = dsPAYMENT.Tables("OINV")
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
            dtINV1 = dsPAYMENT.Tables("INV1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtINV1)

            '------RIN HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"", ""Comments"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ORIN"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '14') "
            dtORIN = dsPAYMENT.Tables("ORIN")
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
            dtRIN1 = dsPAYMENT.Tables("RIN1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRIN1)

            '------PCH HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"", ""Comments"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OPCH"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '18') "
            dtOPCH = dsPAYMENT.Tables("OPCH")
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
            dtPCH1 = dsPAYMENT.Tables("PCH1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPCH1)

            '------RPC HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"", ""Comments"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ORPC"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
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
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '19') "
            dtRPC1 = dsPAYMENT.Tables("RPC1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtRPC1)

            '------DPI HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"", ""Comments"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ODPI"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
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
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '203') "
            dtDPI1 = dsPAYMENT.Tables("DPI1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtDPI1)

            '------DPO HEADER--------------------------------------------------
            sQuery = " SELECT ""DocEntry"",""DocNum"",""Series"",""CardCode"",""DocDate"",""DocDueDate"",""DocCur"",""DocRate"",""DocType"",""DocTotal"",""DocTotalFC"",""NumAtCard"",""ObjType"",""TaxDate"",""VatSum"",""VatSumFC"", ""Comments"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""ODPO"" "
            sQuery &= " WHERE ""DocEntry"" IN ( "
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
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
            sQuery &= " SELECT DISTINCT ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""VPM2"" "
            sQuery &= " WHERE ""DocNum"" = '" & g_sDocEntry & "' AND ""InvType"" = '204') "
            dtDPO1 = dsPAYMENT.Tables("DPO1")
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

            dtOJDT = dsPAYMENT.Tables("OJDT")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOJDT)
            '--------------------------------------------------------
            sQuery = "  SELECT T1.*, IFNULL(T2.""CardName"",'') AS ""OrigCardName"",  T2.""BankCode"", T3.""BankName"", T4.""INTERNAL_K"", T4.""U_NAME"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OVPM"" T1 "
            sQuery &= " LEFT OUTER JOIN  """ & oCompany.CompanyDB & """.""OCRD"" T2 ON T1.""CardCode"" = T2.""CardCode"" "
            sQuery &= " LEFT OUTER JOIN  """ & oCompany.CompanyDB & """.""ODSC"" T3 ON T2.""BankCode"" = T3.""BankCode"" "
            sQuery &= " LEFT OUTER JOIN  """ & oCompany.CompanyDB & """.""OUSR"" T4 ON T1.""UserSign"" = T4.""INTERNAL_K"" "
            sQuery &= " WHERE T1.""DocNum"" = '" & g_sDocNum & "' AND T1.""Series"" = '" & g_sSeries & "' AND T1.""DocEntry"" = '" & g_sDocEntry & "' "
            dtOVPM = dsPAYMENT.Tables("OVPM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOVPM)

            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM1"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM1 = dsPAYMENT.Tables("VPM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM1)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM2"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM2 = dsPAYMENT.Tables("VPM2")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM2)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM3"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM3 = dsPAYMENT.Tables("VPM3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM3)
            '--------------------------------------------------------
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""VPM4"" WHERE ""DocNum"" = '" & g_sDocEntry & "' "
            dtVPM4 = dsPAYMENT.Tables("VPM4")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVPM4)

            '--------------------------------------------------------
            sQuery = " SELECT '1' ""FLAG"", '1' ""SRNO"" FROM DUMMY "
            dtIMAGE = dsPAYMENT.Tables("@NCM_IMAGE")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtIMAGE)

            '--------------------------------------------------------
            sQuery = " SELECT ""AcctCode"",""AcctName"",""LogInstanc"",""FormatCode"" FROM """ & oCompany.CompanyDB & """.""OACT"" "
            dtOACT = dsPAYMENT.Tables("OACT")
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
            dtNNM1 = dsPAYMENT.Tables("NNM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtNNM1)


            '--------------------------------------------------------
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"", ""StreetF"", ""BlockF"", ""ZipCodeF"", ""BuildingF"" FROM """ & oCompany.CompanyDB & """.""ADM1""  "
            dtADM1 = dsPAYMENT.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)

            '--------------------------------------------------------
            sQuery = "SELECT ""FaxF"",""Phone1F"",""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"",""DdctOffice"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
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
    Friend Sub LoadViewer()
        Try
            Dim frm As Hydac_FormViewer = New Hydac_FormViewer
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\OP_{1}.pdf"
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
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\OP\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\OP_" & sCurrDate & sCurrTime & "_" & g_sDocNum & ".pdf"
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
                    Select Case g_sReportType
                        Case "PV"
                            .Text = "Payment Voucher - " & g_sDocNum
                            .ReportName = ReportName.PV
                        Case "RA"
                            .Text = "Remittance Advice - " & g_sDocNum
                            .ReportName = ReportName.RA
                    End Select

                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            .ClientType = "D"
                        Case SAPbouiCOM.BoClientType.ct_Browser
                            .ClientType = "S"
                    End Select

                    .ExportPath = sFinalFileName
                    .Dataset = dsPAYMENT
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
                        frm.OPEN_HANADS_PAYMENTVOUCHER()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select

            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PaymentVoucher].[LoadViewer]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                        If (IsTriggerPV()) Then
                            BubbleEvent = False
                            Dim oDB As SAPbouiCOM.DBDataSource = oForm.DataSources.DBDataSources.Item("OVPM")
                            g_sDocEntry = ""
                            g_sDocEntry = oDB.GetValue("DocEntry", oDB.Offset).ToString.Trim

                            oCombo = oForm.Items.Item("87").Specific
                            g_sSeries = oCombo.Selected.Value
                            oEdit = oForm.Items.Item("3").Specific
                            g_sDocNum = oEdit.Value
                            Dim oRecordset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordset.DoQuery("SELECT ""DocType"", ""CashSum"", ""TrsfrSum"" FROM """ & oCompany.CompanyDB & """.""OVPM"" WHERE ""Series"" ='" & g_sSeries & "' and ""DocNum"" ='" & g_sDocNum & "' AND ""DocEntry"" = '" & g_sDocEntry & "'")
                            If oRecordset.RecordCount > 0 Then
                                Dim bInPV As Boolean = False
                                Dim bInRA As Boolean = False

                                bInPV = SubMain.IsIncludeModule(ReportName.PV)
                                bInRA = SubMain.IsIncludeModule(ReportName.RA)
                                If (bInPV AndAlso bInRA) Then
                                    If SBO_Application.MessageBox("Please select document to view: 1) PV - Payment Voucher 2) RA - Remittance Advice.", 1, "PV", "RA") = 1 Then
                                        g_sReportType = "PV"
                                    Else
                                        g_sReportType = "RA"
                                    End If
                                Else
                                    If (bInPV) Then
                                        g_sReportType = "PV"
                                    ElseIf (bInRA) Then
                                        g_sReportType = "RA"
                                    Else
                                        Return True
                                    End If
                                End If
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
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[PaymentVoucher].[MenuEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
End Class