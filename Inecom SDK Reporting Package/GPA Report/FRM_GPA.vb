Imports System.Threading
Imports SAPbobsCOM

Public Class FRM_GPA

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form 'DW
    Private oRecordset As SAPbobsCOM.Recordset
    Private oItem As SAPbouiCOM.Item
    Private oEdit As SAPbouiCOM.EditText
    Private oMatrix As SAPbouiCOM.Matrix
    Private oColumn As SAPbouiCOM.Column
    Private ds As DataSet
    Private dt, dtProj As DataTable
    Private objDataRow As DataRow
    Private sQuery As String
    Private i As Integer
    Private g_sReportFileName As String = ""
#End Region

#Region "Constructors"
    Public Sub New()
        MyBase.new()
    End Sub
#End Region

#Region "Setting Form"
    Public Sub LoadForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_GPA.srf") = True Then
            oForm = SBO_Application.Forms.Item("NCM_GPA")
            oForm.SupportedModes = -1
            DefineUserDataSource()
            oForm.Visible = True
        Else
            Try
                oForm = SBO_Application.Forms.Item("NCM_GPA")
                If oForm.Visible Then
                    oForm.Select()
                Else
                    oForm.Close()
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub
    Private Sub DefineUserDataSource()
        Try
            oForm.DataSources.UserDataSources.Add("txtFrDate", SAPbouiCOM.BoDataType.dt_DATE)
            oEdit = oForm.Items.Item("txtFrDate").Specific
            oEdit.DataBind.SetBound(True, , "txtFrDate")
            oForm.DataSources.UserDataSources.Add("txtToDate", SAPbouiCOM.BoDataType.dt_DATE)
            oEdit = oForm.Items.Item("txtToDate").Specific
            oEdit.DataBind.SetBound(True, , "txtToDate")
            oForm.DataSources.UserDataSources.Add("txtFrBP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 15)
            oEdit = oForm.Items.Item("txtFrBP").Specific
            oEdit.DataBind.SetBound(True, , "txtFrBP")
            oForm.DataSources.UserDataSources.Add("txtToBP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 15)
            oEdit = oForm.Items.Item("txtToBP").Specific
            oEdit.DataBind.SetBound(True, , "txtToBP")
            oForm.DataSources.UserDataSources.Add("txtFrProj", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            oEdit = oForm.Items.Item("txtFrProj").Specific
            oEdit.DataBind.SetBound(True, , "txtFrProj")
            oForm.DataSources.UserDataSources.Add("txtToProj", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
            oEdit = oForm.Items.Item("txtToProj").Specific
            oEdit.DataBind.SetBound(True, , "txtToProj")
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[DefineDataSource] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "PrintReport"
    Private Function Validate() As Boolean
        Try
            Dim oFrDate, oToDate As SAPbouiCOM.EditText
            oRecordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            oFrDate = oForm.Items.Item("txtFrDate").Specific
            oToDate = oForm.Items.Item("txtToDate").Specific
            If oFrDate.Value.Length > 0 And oToDate.Value.Length > 0 Then
                If oToDate.Value < oFrDate.Value Then
                    Throw New Exception("To date cannot ealier than From date")
                End If
            End If

            If oForm.Items.Item("txtFrBP").Specific.value <> String.Empty Then
                'HANA
                'oRecordset.DoQuery("select 1 from ocrd where cardtype = 'C' and cardcode = '" & oForm.Items.Item("txtFrBP").Specific.value & "'")
                oRecordset.DoQuery("SELECT 1 FROM " & oCompany.CompanyDB & ".OCRD WHERE ""CardType"" = 'C' AND ""CardCode"" = '" & oForm.Items.Item("txtFrBP").Specific.value & "'")
                If oRecordset.RecordCount <= 0 Then
                    Throw New Exception("Invalid From BP Code: " & oForm.Items.Item("txtFrBP").Specific.value)
                End If
            End If

            If oForm.Items.Item("txtToBP").Specific.value <> String.Empty Then
                ' oRecordset.DoQuery("select 1 from ocrd where cardtype = 'C' and cardcode = '" & oForm.Items.Item("txtToBP").Specific.value & "'")
                oRecordset.DoQuery("SELECT 1 FROM " & oCompany.CompanyDB & ".OCRD WHERE ""CardType"" = 'C' AND ""CardCode"" = '" & oForm.Items.Item("txtToBP").Specific.value & "'")
                If oRecordset.RecordCount <= 0 Then
                    Throw New Exception("Invalid To BP Code: " & oForm.Items.Item("txtToBP").Specific.value)
                End If
            End If

            If oForm.Items.Item("txtFrProj").Specific.value <> String.Empty Then
                ' HANA
                'oRecordset.DoQuery("select 1 from oprj where prjcode = '" & oForm.Items.Item("txtFrProj").Specific.value & "'")
                oRecordset.DoQuery("SELECT 1 FROM " & oCompany.CompanyDB & ".OPRJ WHERE ""PrjCode"" = '" & oForm.Items.Item("txtFrProj").Specific.value & "'")
                If oRecordset.RecordCount <= 0 Then
                    Throw New Exception("Invalid From Project Code: " & oForm.Items.Item("txtFrProj").Specific.value)
                End If
            End If

            If oForm.Items.Item("txtToProj").Specific.value <> String.Empty Then
                'oRecordset.DoQuery("select 1 from oprj where prjcode = '" & oForm.Items.Item("txtToProj").Specific.value & "'")
                oRecordset.DoQuery("SELECT 1 FROM " & oCompany.CompanyDB & ".OPRJ WHERE ""PrjCode"" = '" & oForm.Items.Item("txtToProj").Specific.value & "'")
                If oRecordset.RecordCount <= 0 Then
                    Throw New Exception("Invalid To Project Code: " & oForm.Items.Item("txtToProj").Specific.value)
                End If
            End If

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFileName = GetSharedFilePath(ReportName.GPA)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[GPA].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        Try
            ds = New dsRpt
            dt = ds.Tables("TMP_GPA")
            dtProj = ds.Tables("TMP_PROJ")
            LoadReport()

            Dim frm As New Hydac_FormViewer
            frm.IsShared = IsSharedFileExist()
            frm.SharedReportName = g_sReportFileName
            With oForm.DataSources.UserDataSources
                If .Item("txtFrDate").ValueEx <> "" Then
                    frm.FromDate = .Item("txtFrDate").ValueEx
                End If
                If .Item("txtToDate").ValueEx <> "" Then
                    frm.ToDate = .Item("txtToDate").ValueEx
                End If
                If .Item("txtFrBP").ValueEx <> "" Then
                    frm.FromBP = .Item("txtFrBP").ValueEx
                End If
                If .Item("txtToBP").ValueEx <> "" Then
                    frm.ToBP = .Item("txtToBP").ValueEx
                End If
                If .Item("txtFrProj").ValueEx <> "" Then
                    frm.FromProj = .Item("txtFrProj").ValueEx
                End If
                If .Item("txtToProj").ValueEx <> "" Then
                    frm.ToProj = .Item("txtToProj").ValueEx
                End If
            End With

            frm.ReportName = ReportName.GPA
            frm.Dataset = ds
            frm.ShowDialog()
        Catch ex As Exception
            SBO_Application.MessageBox("[LoadViewer] : " & ex.Message)
        End Try
    End Sub
    Private Sub LoadReport()
        Dim oRecordset1 As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
        Try
            Dim oFrDate, oToDate As SAPbouiCOM.EditText
            Dim qryFrDate As String = String.Empty
            Dim qryToDate As String = String.Empty
            Dim qryToBP As String = String.Empty
            Dim qryToProj As String = String.Empty
            Dim dr As System.Data.DataRow
            Dim objRows As System.Data.DataRow()

            dt.Clear()
            dtProj.Clear()

            'Load Data to temp table
            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oFrDate = oForm.Items.Item("txtFrDate").Specific
            oToDate = oForm.Items.Item("txtToDate").Specific

            If oFrDate.Value.Length > 0 Then
                ' HANA
                'qryFrDate = "and i.docdate >= '" & oFrDate.Value & "' "
                qryFrDate = "and i.""DocDate"" >= '" & oFrDate.Value & "' "
            End If
            If oToDate.Value.Length > 0 Then
                'HANA
                'qryToDate = "and i.docdate <= '" & oToDate.Value & "' "
                qryToDate = "and i.""DocDate"" <= '" & oToDate.Value & "' "
            End If
            If oForm.Items.Item("txtToBP").Specific.Value.Length > 0 Then
                'HANA
                'qryToBP = "and i.cardcode <= '" & oForm.Items.Item("txtToBP").Specific.Value & "' "
                qryToBP = "and i.""CardCode"" <= '" & oForm.Items.Item("txtToBP").Specific.Value & "' "
            End If
            If oForm.Items.Item("txtToProj").Specific.Value.Length > 0 Then
                ' HANA
                'qryToProj = "and i.project <= '" & oForm.Items.Item("txtToProj").Specific.Value & "' "
                qryToProj = "and i.""Project"" <= '" & oForm.Items.Item("txtToProj").Specific.Value & "' "
            End If

            ' HANA
            'sQuery = " Select isnull(i.project,'') as project, p.prjname, (i.cardcode + ' - ' + i.cardname) as 'Cust', " & _
            '        "  (CASE WHEN n1.seriesname LIKE 'ARDN%' THEN 'DN' ELSE 'IN' END) as 'DocType', " & _
            '        "  i.docnum, i.docentry, i.docdate, i.doccur, sum(i1.linetotal) as 'LCAmt', sum(i1.totalfrgn) as 'FCAmt' " & _
            '        "  From OINV i " & _
            '        "  LEFT OUTER JOIN INV1 i1 ON i.docentry = i1.docentry " & _
            '        "  LEFT OUTER JOIN OPRJ p  ON i.project = p.prjcode " & _
            '        "  LEFT OUTER JOIN NNM1 n1 ON n1.series = i.series AND n1.objectcode = 13 " & _
            '        "  Where i.cardcode >= '" & oForm.Items.Item("txtFrBP").Specific.Value & "' "
            sQuery = " SELECT IFNULL(i.""Project"", '') AS ""project"", p.""PrjName"", (i.""CardCode"" || ' - ' || i.""CardName"") AS ""Cust"", ( " & _
                    "   CASE WHEN n1.""SeriesName"" LIKE 'ARDN%' THEN 'DN' " & _
                    "          ELSE 'IN' END) AS ""DocType"", " & _
                    "   i.""DocNum"", i.""DocEntry"", i.""DocDate"", i.""DocCur"", SUM(i1.""LineTotal"") AS ""LCAmt"",  " & _
                    "   SUM(i1.""TotalFrgn"") AS ""FCAmt""  " & _
                    "  FROM OINV i  " & _
                    "       LEFT OUTER JOIN INV1 i1 ON i.""DocEntry"" = i1.""DocEntry""   " & _
                    "       LEFT OUTER JOIN OPRJ p ON i.""Project"" = p.""PrjCode""  " & _
                    "       LEFT OUTER JOIN NNM1 n1 ON n1.""Series"" = i.""Series"" AND n1.""ObjectCode"" = 13  " & _
                    "  Where i.""CardCode"" >= '" & oForm.Items.Item("txtFrBP").Specific.Value & "' "

            sQuery &= qryToBP
            sQuery &= qryFrDate
            sQuery &= qryToDate

            'HANA
            'sQuery &= "and i.project >= '" & oForm.Items.Item("txtFrProj").Specific.Value & "' "
            sQuery &= "and i.""Project"" >= '" & oForm.Items.Item("txtFrProj").Specific.Value & "' "

            sQuery &= qryToProj

            'HANA
            'sQuery &= "and isnull(i.project,'') in (select isnull(i2.project,'') as project from pch1 i2 " & _
            '        " union select isnull(v4.project,'') as project from ovpm v left join vpm4 v4 on v.docentry = v4.docnum " & _
            '        " where v.doctype = 'A' " & _
            '        " union select isnull(r.project,'') as project from rpc1 r) " & _
            '        " group by i.project, p.prjname, i.cardcode, i.cardname, i.docnum, n1.seriesname, i.docentry, i.docdate, i.doccur " & _
            '        " UNION " & _
            '        " Select isnull(i.project,'') as project, p.prjname, (i.cardcode + ' - ' + i.cardname) as 'Cust', " & _
            '        " 'CN' as 'DocType', i.docnum, i.docentry, i.docdate, i.doccur, " & _
            '        " sum(i1.linetotal) * -1 as 'LCAmt', sum(i1.totalfrgn) * -1 as 'FCAmt' " & _
            '        " From ORIN i " & _
            '        " LEFT OUTER JOIN RIN1 i1 ON i.docentry = i1.docentry " & _
            '        " LEFT OUTER JOIN OPRJ p ON i.project = p.prjcode " & _
            '        " Where i.cardcode >= '" & oForm.Items.Item("txtFrBP").Specific.Value & "' "
            sQuery &= "  AND  IFNULL(i.""Project"", '') IN (SELECT IFNULL(i2.""Project"", '') AS ""project"" " & _
                                "   FROM PCH1 i2  " & _
                                " UNION  " & _
                                "     SELECT IFNULL(v4.""Project"", '') AS ""project""   " & _
                                "     FROM OVPM v   " & _
                                "         LEFT OUTER JOIN VPM4 v4 ON v.""DocEntry"" = v4.""DocNum""   " & _
                                "     WHERE v.""DocType"" = 'A'   " & _
                                "  UNION " & _
                                "      SELECT IFNULL(r.""Project"", '') AS ""project""  " & _
                                "      FROM RPC1 r)  " & _
                                "  GROUP BY i.""Project"", p.""PrjName"", i.""CardCode"", i.""CardName"", i.""DocNum"", n1.""SeriesName"", i.""DocEntry"", i.""DocDate"", " & _
                                "       i.""DocCur""  " & _
                                "  UNION " & _
                                "  SELECT IFNULL(i.""Project"", '') AS ""project"", p.""PrjName"", (i.""CardCode"" || ' - ' || i.""CardName"") AS ""Cust"",  " & _
                                "      'CN' AS ""DocType"", i.""DocNum"", i.""DocEntry"", i.""DocDate"", i.""DocCur"", SUM(i1.""LineTotal"") * -1 AS ""LCAmt"",  " & _
                                "      SUM(i1.""TotalFrgn"") * -1 AS ""FCAmt""  " & _
                                "  FROM ORIN i  " & _
                                "      LEFT OUTER JOIN RIN1 i1 ON i.""DocEntry"" = i1.""DocEntry""  " & _
                                "      LEFT OUTER JOIN OPRJ p ON i.""Project"" = p.""PrjCode""  " & _
                                " Where i.""CardCode"" >= '" & oForm.Items.Item("txtFrBP").Specific.Value & "' "

            sQuery &= qryToBP
            sQuery &= qryFrDate
            sQuery &= qryToDate

            'HANA
            'sQuery &= "and i.project >= '" & oForm.Items.Item("txtFrProj").Specific.Value & "' "
            sQuery &= "and i.""Project"" >= '" & oForm.Items.Item("txtFrProj").Specific.Value & "' "

            sQuery &= qryToProj

            ' HANA
            'sQuery &= "and isnull(i.project,'') in (select isnull(i2.project,'') as project from pch1 i2 " & _
            '        " union select isnull(v4.project,'') as project from ovpm v left join vpm4 v4 on v.docentry = v4.docnum " & _
            '        " where v.doctype = 'A' " & _
            '        " union select isnull(r.project,'') as project from rpc1 r) " & _
            '        " group by i.project, p.prjname, i.cardcode, i.cardname, i.docnum, i.docentry, i.docdate, i.doccur " & _
            '        " order by project, i.docdate, doctype desc"

            sQuery &= " and IFNULL(i.""Project"", '') IN (SELECT IFNULL(i2.""Project"", '') AS ""project"" " & _
                   "     FROM PCH1 i2   " & _
                   " UNION " & _
                   "     SELECT IFNULL(v4.""Project"", '') AS ""project""  " & _
                   "     FROM OVPM v  " & _
                   "         LEFT OUTER JOIN VPM4 v4 ON v.""DocEntry"" = v4.""DocNum""  " & _
                   "     WHERE v.""DocType"" = 'A'  " & _
                   " UNION " & _
                   "     SELECT IFNULL(r.""Project"", '') AS ""project""  " & _
                   "     FROM RPC1 r)  " & _
                   " GROUP BY i.""Project"", p.""PrjName"", i.""CardCode"", i.""CardName"", i.""DocNum"", i.""DocEntry"", i.""DocDate"", i.""DocCur""  " & _
                   " ORDER BY ""project"", i.""DocDate"", ""DocType"" DESC "

            oRecordset.DoQuery(sQuery)
            If oRecordset.RecordCount > 0 Then
                Do Until oRecordset.EoF
                    SBO_Application.StatusBar.SetText("Reading AR Document No... " & oRecordset.Fields.Item("DocNum").Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objDataRow = dt.NewRow
                    objDataRow("Project") = oRecordset.Fields.Item("Project").Value
                    objDataRow("PrjName") = oRecordset.Fields.Item("PrjName").Value
                    objDataRow("DocType") = oRecordset.Fields.Item("DocType").Value
                    objDataRow("DocNum") = oRecordset.Fields.Item("DocNum").Value
                    objDataRow("DocDate") = oRecordset.Fields.Item("DocDate").Value
                    objDataRow("Curr") = oRecordset.Fields.Item("DocCur").Value
                    objDataRow("LCAmt") = oRecordset.Fields.Item("LCAmt").Value
                    objDataRow("FCAmt") = oRecordset.Fields.Item("FCAmt").Value
                    objDataRow("Cust") = oRecordset.Fields.Item("Cust").Value
                    objDataRow("APDocType") = ""
                    objDataRow("GProfitLC") = objDataRow("LCAmt") - 0
                    If (objDataRow("GProfitLC") <> 0 And objDataRow("LCAmt") <> 0) Then
                        objDataRow("GProfitPerc") = (objDataRow("GProfitLC") / objDataRow("LCAmt")) * 100
                    Else
                        objDataRow("GProfitPerc") = 0
                    End If
                    dt.Rows.Add(objDataRow)

                    objRows = dtProj.Select("Project = '" & oRecordset.Fields.Item("project").Value & "'")
                    If objRows.GetUpperBound(0) < 0 Then
                        dr = dtProj.NewRow
                        dr("Project") = oRecordset.Fields.Item("Project").Value
                        dtProj.Rows.Add(dr)
                    End If

                    oRecordset.MoveNext()
                Loop
            Else
                SBO_Application.StatusBar.SetText("No AR records found.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

            'Get AP Document
            objRows = dtProj.Select()
            If objRows.GetUpperBound(0) >= 0 Then
                For Each dr In objRows

                    ' HANA
                    'sQuery = "Select isnull(i1.project,'') as 'Project', p.prjname, 'IN' as 'DocType', i.docnum, i.docentry, " & _
                    '        "i.docdate as docdate, i.doccur as 'DocCur', " & _
                    '        "sum(i1.linetotal) as 'LCAmt', sum(i1.totalfrgn) as 'FCAmt' " & _
                    '        "From OPCH i Left Join PCH1 i1 ON i.docentry = i1.docentry " & _
                    '        "Left Join OPRJ p ON i.project = p.prjcode " & _
                    '        "where isnull(i1.project,'') = '" & dr("Project") & "' " & _
                    '        "and isnull(i1.project,'') in (select isnull(i2.project,'') as project from oinv i2 " & _
                    '        "union select isnull(r1.project,'') as project from orin r1) " & _
                    '        "group by i1.project, p.prjname, i.docnum, i.docentry, i.docdate, i.doccur " & _
                    '        "UNION " & _
                    '        "Select isnull(i1.project,'') as 'Project', p.prjname, 'CN' as 'DocType', i.docnum, i.docentry, " & _
                    '        "i.docdate as docdate, i.doccur as 'DocCur', " & _
                    '        "sum(i1.linetotal) * -1 as 'LCAmt', sum(i1.totalfrgn) * -1 as 'FCAmt' " & _
                    '        "From ORPC i Left Join RPC1 i1 ON i.docentry = i1.docentry " & _
                    '        "Left Join OPRJ p ON i.project = p.prjcode " & _
                    '        "where isnull(i1.project,'') = '" & dr("Project") & "' " & _
                    '        "and isnull(i1.project,'') in (select isnull(i2.project,'') as project from oinv i2 " & _
                    '        "union select isnull(r1.project,'') as project from orin r1) " & _
                    '        "group by i1.project, p.prjname, i.docnum, i.docentry, i.docdate, i.doccur " & _
                    '        "UNION " & _
                    '        "Select isnull(v4.project,'') as 'Project', p.prjname, 'PY' as 'DocType', v.docnum, v.docentry, " & _
                    '        "v.docdate as docdate, v.doccurr as 'DocCur', v4.sumapplied as 'LCAmt', v4.appliedfc as 'FCAmt' " & _
                    '        "from ovpm v Left join vpm4 v4 on v.docentry = v4.docnum " & _
                    '        "Left Join OPRJ p ON v4.project = p.prjcode " & _
                    '        "where v.doctype = 'A' " & _
                    '        "and isnull(v4.project,'') = '" & dr("Project") & "' " & _
                    '        "and isnull(v4.project,'') in (select isnull(i2.project,'') as project from oinv i2 " & _
                    '        "union select isnull(r1.project,'') as project from orin r1) " & _
                    '        "order by project, docdate, doctype desc"

                    sQuery = " SELECT IFNULL(i1.""Project"", '') AS ""Project"", p.""PrjName"", 'IN' AS ""DocType"", i.""DocNum"", i.""DocEntry"",  " & _
                                "     i.""DocDate"" AS ""docdate"", i.""DocCur"" AS ""DocCur"", SUM(i1.""LineTotal"") AS ""LCAmt"", SUM(i1.""TotalFrgn"") AS ""FCAmt""  " & _
                                " FROM OPCH i  " & _
                                "     LEFT OUTER JOIN PCH1 i1 ON i.""DocEntry"" = i1.""DocEntry""  " & _
                                "     LEFT OUTER JOIN OPRJ p ON i.""Project"" = p.""PrjCode""  " & _
                                " WHERE IFNULL(i1.""Project"", '') = '" & dr("Project") & "' " & _
                                "   AND IFNULL(i1.""Project"", '') IN (SELECT IFNULL(i2.""Project"", '') AS ""project""  " & _
                                "       FROM OINV i2  " & _
                                "       UNION  " & _
                                "     SELECT IFNULL(r1.""Project"", '') AS ""project""  " & _
                                "     FROM ORIN r1)  " & _
                                " GROUP BY i1.""Project"", p.""PrjName"", i.""DocNum"", i.""DocEntry"", i.""DocDate"", i.""DocCur""  " & _
                                "       UNION  " & _
                                 " SELECT IFNULL(i1.""Project"", '') AS ""Project"", p.""PrjName"", 'CN' AS ""DocType"", i.""DocNum"", i.""DocEntry"",  " & _
                                 "     i.""DocDate"" AS ""docdate"", i.""DocCur"" AS ""DocCur"", SUM(i1.""LineTotal"") * -1 AS ""LCAmt"",  " & _
                                 "     SUM(i1.""TotalFrgn"") * -1 AS ""FCAmt""  " & _
                                 "     FROM ORPC i  " & _
                                 "      LEFT OUTER JOIN RPC1 i1 ON i.""DocEntry"" = i1.""DocEntry"" " & _
                                 "     LEFT OUTER JOIN OPRJ p ON i.""Project"" = p.""PrjCode""  " & _
                                " WHERE IFNULL(i1.""Project"", '') = '" & dr("Project") & "' " & _
                                 " AND IFNULL(i1.""Project"", '') IN (SELECT IFNULL(i2.""Project"", '') AS ""project""  " & _
                                 "     FROM OINV i2  " & _
                                 " UNION  " & _
                                 "     SELECT IFNULL(r1.""Project"", '') AS ""project""  " & _
                                 "     FROM ORIN r1)  " & _
                                  " GROUP BY i1.""Project"", p.""PrjName"", i.""DocNum"", i.""DocEntry"", i.""DocDate"", i.""DocCur""  " & _
                                 " UNION  " & _
                                  " SELECT IFNULL(v4.""Project"", '') AS ""Project"", p.""PrjName"", 'PY' AS ""DocType"", v.""DocNum"", v.""DocEntry"",  " & _
                                 "     v.""DocDate"" AS ""docdate"", v.""DocCurr"" AS ""DocCur"", v4.""SumApplied"" AS ""LCAmt"", v4.""AppliedFC"" AS ""FCAmt""  " & _
                                  " FROM OVPM v  " & _
                                 "     LEFT OUTER JOIN VPM4 v4 ON v.""DocEntry"" = v4.""DocNum""  " & _
                                  "     LEFT OUTER JOIN OPRJ p ON v4.""Project"" = p.""PrjCode""  " & _
                                 " WHERE v.""DocType"" = 'A' AND  " & _
                                " IFNULL(v4.""Project"", '') = '" & dr("Project") & "' " & _
                                "        AND  IFNULL(v4.""Project"", '') IN (SELECT IFNULL(i2.""Project"", '') AS ""project""  " & _
                                "     FROM OINV i2  " & _
                                 " UNION  " & _
                                "     SELECT IFNULL(r1.""Project"", '') AS ""project""  " & _
                                 "     FROM ORIN r1)  " & _
                                    " ORDER BY ""Project"", ""DocDate"", ""DocType"" DESC "

                    oRecordset.DoQuery(sQuery)
                    Do Until oRecordset.EoF
                        SBO_Application.StatusBar.SetText("Reading AP Document No... " & oRecordset.Fields.Item("DocNum").Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        Dim Rows As System.Data.DataRow() = dt.Select("Project = '" & oRecordset.Fields.Item("Project").Value & "' and APDocType = ''")
                        If Rows.GetUpperBound(0) >= 0 Then
                            For Each objDataRow In Rows

                                objDataRow.BeginEdit()
                                objDataRow.Item("APDocNum") = oRecordset.Fields.Item("DocNum").Value
                                objDataRow.Item("APDocType") = oRecordset.Fields.Item("DocType").Value
                                objDataRow.Item("APDocDate") = oRecordset.Fields.Item("DocDate").Value
                                objDataRow.Item("APCurr") = oRecordset.Fields.Item("DocCur").Value
                                objDataRow.Item("APFCAmt") = oRecordset.Fields.Item("FCAmt").Value
                                objDataRow("APLCAmt") = oRecordset.Fields.Item("LCAmt").Value
                                objDataRow("GProfitLC") = objDataRow("LCAmt") - objDataRow("APLCAmt")
                                If (objDataRow("GProfitLC") <> 0 And objDataRow("LCAmt") <> 0) Then
                                    objDataRow("GProfitPerc") = (objDataRow("GProfitLC") / objDataRow("LCAmt")) * 100
                                Else
                                    objDataRow("GProfitPerc") = 0
                                End If
                                objDataRow.EndEdit()
                                objDataRow.AcceptChanges()
                                Exit For
                            Next
                        Else
                            objDataRow = dt.NewRow
                            objDataRow("Project") = oRecordset.Fields.Item("Project").Value
                            objDataRow("PrjName") = oRecordset.Fields.Item("PrjName").Value
                            objDataRow("APDocType") = oRecordset.Fields.Item("DocType").Value
                            objDataRow("APDocNum") = oRecordset.Fields.Item("DocNum").Value
                            objDataRow("APDocDate") = oRecordset.Fields.Item("DocDate").Value
                            objDataRow("APCurr") = oRecordset.Fields.Item("DocCur").Value
                            objDataRow("APLCAmt") = oRecordset.Fields.Item("LCAmt").Value
                            objDataRow("APFCAmt") = oRecordset.Fields.Item("FCAmt").Value
                            objDataRow("GProfitLC") = 0 - objDataRow("APLCAmt")
                            objDataRow("GProfitPerc") = 0
                            dt.Rows.Add(objDataRow)

                        End If
                        oRecordset.MoveNext()
                    Loop
                Next
            End If

            oRecordset = Nothing
            oRecordset1 = Nothing
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                    Case "txtFrBP", "txtToBP", "txtFrProj", "txtToProj"
                        If pval.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN AndAlso pval.CharPressed = keyID.Tab Then
                            oEdit = oForm.Items.Item(pval.ItemUID).Specific
                            If oEdit.Value = String.Empty Then
                                BubbleEvent = False
                                SBO_Application.SendKeys("+{F2}")
                            End If
                        End If
                    Case Else
                        ' do nothing
                End Select
            Else
                If pval.ItemUID = "btnPrint" AndAlso pval.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If Validate() Then
                        Dim myThread As Thread = New Thread(New ThreadStart(AddressOf LoadViewer))
                        myThread.SetApartmentState(ApartmentState.STA)
                        myThread.Start()
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
