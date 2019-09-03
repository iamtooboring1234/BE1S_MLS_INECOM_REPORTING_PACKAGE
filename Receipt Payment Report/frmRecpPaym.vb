Imports System.IO
Imports System.Xml
Imports System.Data.Common

Public Class frmRecpPaym

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oItem As SAPbouiCOM.Item
    Private oStatic As SAPbouiCOM.StaticText
    Private g_sReportFilename As String = String.Empty
    Private dt As System.Data.DataTable
    Private dt_DETAILS As System.Data.DataTable
    Private ds As System.Data.DataSet

    Dim g_bIsShared As Boolean = False
    Dim oCheck As SAPbouiCOM.CheckBox
    Dim g_iSecond As Integer = 0

#End Region

#Region "Initialisation"
    Public Sub LoadForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmRecpPaymList.srf") Then
            oForm = SBO_Application.Forms.Item("NCM_RecpPaymList")
            AddDataSource()
            If (Not oForm.Visible) Then
                oForm.Visible = True
            End If
            SetupChooseFromList()
            oForm.Items.Item("dtPostFrom").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Else
            Try
                oForm = SBO_Application.Forms.Item("FRM_NCM_RECPPAYM")
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
            Dim oCBox As SAPbouiCOM.ComboBox

            'pick the fields from screen painter.
            oForm.DataSources.UserDataSources.Add("dtPostFrom", SAPbouiCOM.BoDataType.dt_DATE, 254)
            oForm.DataSources.UserDataSources.Add("dtPostTo", SAPbouiCOM.BoDataType.dt_DATE, 1)
            oForm.DataSources.UserDataSources.Add("txtORCTFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oForm.DataSources.UserDataSources.Add("txtORCTTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oForm.DataSources.UserDataSources.Add("txtOVPMFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oForm.DataSources.UserDataSources.Add("txtOVPMTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            oForm.DataSources.UserDataSources.Add("cbType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)

            oCBox = oForm.Items.Item("cbType").Specific
            oCBox.DataBind.SetBound(True, String.Empty, "cbType")
            oCBox.ValidValues.Add("-", "All")
            oCBox.ValidValues.Add("24", "Incoming Payment")
            oCBox.ValidValues.Add("46", "Outgoing Payment")
            oCBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

            oEdit = DirectCast(oForm.Items.Item("dtPostFrom").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "dtPostFrom")
            oForm.DataSources.UserDataSources.Item("dtPostFrom").ValueEx = DateTime.Now.ToString("yyyy0101")

            oEdit = DirectCast(oForm.Items.Item("dtPostTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "dtPostTo")
            Dim sTemp As String = String.Empty
            sTemp = DateTime.Now.Year.ToString("000#") + DateTime.Now.Month.ToString("0#") + DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString("0#")
            oForm.DataSources.UserDataSources.Item("dtPostTo").ValueEx = sTemp

            oEdit = DirectCast(oForm.Items.Item("txtORCTFr").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtORCTFr")

            oEdit = DirectCast(oForm.Items.Item("txtORCTTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtORCTTo")

            oEdit = DirectCast(oForm.Items.Item("txtOVPMFr").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtOVPMFr")

            oEdit = DirectCast(oForm.Items.Item("txtOVPMTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtOVPMTo")

        Catch ex As Exception
            SBO_Application.MessageBox("[frmRecpPaym].[AddDataSource]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams

        Try
            oCFLs = oForm.ChooseFromLists

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "24"
            oCFLCreation.UniqueID = "cflSORCT"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtORCTFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflSORCT"
            oEditLn.ChooseFromListAlias = "DocNum"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "24"
            oCFLCreation.UniqueID = "cflEORCT"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtORCTTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflEORCT"
            oEditLn.ChooseFromListAlias = "DocNum"

            'oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            'oCFLCreation.MultiSelection = False
            'oCFLCreation.ObjectType = "46"
            'oCFLCreation.UniqueID = "cflSOVPM"
            'oCFL = oCFLs.Add(oCFLCreation)

            'oEditLn = DirectCast(oForm.Items.Item("txtOVPMFr").Specific, SAPbouiCOM.EditText)
            'oEditLn.ChooseFromListUID = "cflSOVPM"
            'oEditLn.ChooseFromListAlias = "DocNum"

            'oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            'oCFLCreation.MultiSelection = False
            'oCFLCreation.ObjectType = "46"
            'oCFLCreation.UniqueID = "cflEOVPM"
            'oCFL = oCFLs.Add(oCFLCreation)

            'oEditLn = DirectCast(oForm.Items.Item("txtOVPMTo").Specific, SAPbouiCOM.EditText)
            'oEditLn.ChooseFromListUID = "cflEOVPM"
            'oEditLn.ChooseFromListAlias = "DocNum"

        Catch ex As Exception
            Throw New Exception("[frmRecpPaym].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Sub LoadViewer()
        oForm.Items.Item("btnPrint").Enabled = False
        Dim sFinalExportPath As String = ""
        Dim sFinalFileName As String = ""

        Try
            Dim frm As New RecpPaym_FrmViewer
            Dim bIsContinue As Boolean = False
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\RPL_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Try
                ds = New dsRecpPaym
                dt = ds.Tables("DS_RPT_RECPPAYM")
                dt_DETAILS = ds.Tables("DS_RPT_DETAILS")

                oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
                End If
                oRec = Nothing

                ' ===============================================================================
                ' get the folder of the report of the current DB Name
                ' set to local
                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\RPL\" & oCompany.CompanyDB
                Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                If Not di.Exists Then
                    di.Create()
                End If
                sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                sFinalFileName = di.FullName & "\RPL_" & sCurrDate & "_" & sCurrTime & ".pdf"
                ' ===============================================================================

                If ExecuteProcedure() Then
                    bIsContinue = True

                    Select Case oForm.DataSources.UserDataSources.Item("cbType").ValueEx
                        Case "-"
                            frm.DocumentType = "All"
                        Case "24"
                            frm.DocumentType = "Incoming Payment"
                        Case "46"
                            frm.DocumentType = "Outgoing Payment"
                    End Select

                    frm.ReportDataSet = ds
                    frm.SharedReportName = g_sReportFilename
                    frm.IsReportExternal = g_bIsShared
                    frm.Report = ReportName.RecpPaym
                    frm.Text = "Receipts and Payments Listing"
                    frm.Name = "Receipts and Payments Listing"
                    frm.ExportPath = sFinalFileName
                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            frm.ClientType = "D"
                        Case SAPbouiCOM.BoClientType.ct_Browser
                            frm.ClientType = "S"
                    End Select

                End If
            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item("btnPrint").Enabled = True
            End Try

            If bIsContinue Then
                Select Case SBO_Application.ClientType
                    Case SAPbouiCOM.BoClientType.ct_Desktop
                        frm.ShowDialog()

                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.OpenRPLReport()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[frmRecpPaym].[LoadViewer] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ExecuteProcedure() As Boolean
        g_bIsShared = IsSharedFileExist()

        Dim dbConn As DbConnection = Nothing
        Dim sSORCT As String = String.Empty
        Dim sEORCT As String = String.Empty
        Dim sSOVPM As String = String.Empty
        Dim sEOVPM As String = String.Empty
        Dim sQuery As String = ""
        Dim sSelect As String = ""

        Try
            sSORCT = oForm.Items.Item("txtORCTFr").Specific.value.ToString.Trim
            sEORCT = oForm.Items.Item("txtORCTTo").Specific.value.ToString.Trim
            sSOVPM = oForm.Items.Item("txtOVPMFr").Specific.value.ToString.Trim
            sEOVPM = oForm.Items.Item("txtOVPMTo").Specific.value.ToString.Trim

            If sSORCT.Length <= 0 Then
                sSORCT = "1"
            End If
            If sSOVPM.Length <= 0 Then
                sSOVPM = "1"
            End If
            If sEORCT.Length <= 0 Then
                sEORCT = "9999999999"
            End If
            If sEOVPM.Length <= 0 Then
                sEOVPM = "9999999999"
            End If
            dt.Clear()

            With oForm.DataSources.UserDataSources
                'INCOMING PAYMENT

                SBO_Application.StatusBar.SetText("Collecting Payment Data...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                sQuery = " SELECT X.*, "
                sQuery &= " '" & oForm.Items.Item("dtPostFrom").Specific.value.ToString.Trim & "' AS ""pPostFr"", "
                sQuery &= " '" & oForm.Items.Item("dtPostTo").Specific.value.ToString.Trim & "'   AS ""pPostTo"", "
                sQuery &= " '" & sSORCT & "' AS ""pORCTFr"", "
                sQuery &= " '" & sEORCT & "' AS ""pORCTTo"", "
                sQuery &= " '" & sSOVPM & "' AS ""pOVPMFr"", "
                sQuery &= " '" & sEOVPM & "' AS ""pOVPMTo"" FROM ( "

                sQuery &= " SELECT T0.""ObjType"", CAST(T0.""DocNum"" AS varchar(10)) AS ""DocNumStr"", T0.""DocNum"", T0.""DocDate"",  " & vbCrLf
                sQuery &= " T0.""CounterRef"", T0.""CardCode"", T0.""Canceled"", IFNULL(T2.""Segment_0"", T2.""AcctCode"") || ( " & vbCrLf
                sQuery &= " CASE WHEN T2.""Segment_1"" IS NULL THEN ''  ELSE '-' || T2.""Segment_1"" END) || ( " & vbCrLf
                sQuery &= " CASE WHEN T2.""Segment_2"" IS NULL THEN '' ELSE '-' || T2.""Segment_2"" END) || ( " & vbCrLf
                sQuery &= " CASE WHEN T2.""Segment_3"" IS NULL THEN '' ELSE '-' || T2.""Segment_3"" END) || ( " & vbCrLf
                sQuery &= " CASE WHEN T2.""Segment_4"" IS NULL THEN '' ELSE '-' || T2.""Segment_4"" END) || ( " & vbCrLf
                sQuery &= " CASE WHEN T2.""Segment_5"" IS NULL THEN '' ELSE '-' || T2.""Segment_5"" END) AS ""AcctCode"", " & vbCrLf
                sQuery &= " (CASE WHEN T1.""ShortName"" = T1.""Account"" THEN ''  ELSE T1.""ShortName"" END) AS ""ShortName"", " & vbCrLf
                sQuery &= " T2.""AcctName"", " & vbCrLf
                sQuery &= " (CASE WHEN T1.""Account"" = T1.""ShortName"" THEN ''  " & vbCrLf
                sQuery &= "     ELSE IFNULL( " & vbCrLf
                sQuery &= "             	(SELECT T9.""CardName"" FROM """ & oCompany.CompanyDB & """.OCRD T9 WHERE T9.""CardCode"" = T1.""ShortName""), '')  " & vbCrLf
                sQuery &= "  END) AS ""CardName"", " & vbCrLf
                sQuery &= " T1.""Debit"", T1.""Credit"", T1.""FCCurrency"", T1.""FCDebit"", T1.""FCCredit"",  " & vbCrLf
                sQuery &= " CASE WHEN T2.""Finanse"" = 'Y' THEN 1  ELSE 2  END AS ""BankLine"", ( " & vbCrLf

                ' Cheque(Payment)
                sQuery &= " CASE WHEN T2.""Finanse"" = 'Y'  " & vbCrLf
                sQuery &= " THEN (CASE WHEN (T0.""CheckSum"" <> 0  " & vbCrLf
                sQuery &= "     AND T0.""CheckSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN " & vbCrLf
                sQuery &= " 		'By Cheque : Due on ' " & vbCrLf
                sQuery &= "         || IFNULL( " & vbCrLf
                sQuery &= "             (SELECT  " & vbCrLf
                sQuery &= "             --TOP 1  " & vbCrLf
                sQuery &= "            			CAST(T8.""DueDate"" AS varchar(12)) || '  Cheque ' || CAST(T8.""CheckNum"" AS varchar(10))  " & vbCrLf
                sQuery &= "               		|| '  Amt ' || T8.""Currency"" || ' '  " & vbCrLf
                sQuery &= " 					|| LEFT(CAST(ROUND(T8.""CheckSum"", IFNULL((SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) , LOCATE(T8.""CheckSum"", '.')  " & vbCrLf
                sQuery &= " 							    + IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "             FROM """ & oCompany.CompanyDB & """.RCT1 T8 " & vbCrLf
                sQuery &= "             WHERE T8.""DocNum"" = T0.""DocNum""  " & vbCrLf
                sQuery &= " 				AND T8.""CheckSum"" = ( CASE  WHEN T0.""CheckSumFC"" <> 0 THEN T0.""CheckSumFC""  ELSE T0.""CheckSum"" 	END) " & vbCrLf
                sQuery &= "             ), '') " & vbCrLf
                sQuery &= "     ELSE '' " & vbCrLf
                sQuery &= "     END)  " & vbCrLf

                'Transfer Payment
                sQuery &= " || ( CASE WHEN (T0.""TrsfrSum"" <> 0 " & vbCrLf
                sQuery &= "         AND T0.""TrsfrSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN " & vbCrLf
                sQuery &= " 			'By Bank Transfer : ' || CAST(T0.""TrsfrDate"" AS varchar(12)) || '  ' || IFNULL(T0. ""TrsfrRef"", '') || '  Amt '  " & vbCrLf
                sQuery &= " 			|| (  CASE WHEN T0.""TrsfrSumFC"" <> 0 THEN T0.""DocCurr""  " & vbCrLf
                sQuery &= "                     		ELSE IFNULL((SELECT T9.""MainCurncy"" FROM """ & oCompany.CompanyDB & """.OADM T9), '')  " & vbCrLf
                sQuery &= "                 		END) || ' '  " & vbCrLf
                sQuery &= " 			|| ( CASE WHEN T0.""TrsfrSumFC"" <> 0 THEN  " & vbCrLf
                sQuery &= " 				LEFT(CAST(ROUND(T0.""TrsfrSumFC"", " & vbCrLf
                sQuery &= " 					 IFNULL((SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 					, LOCATE(T0.""TrsfrSumFC"", '.')  " & vbCrLf
                sQuery &= " 				+ IFNULL(  (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "                  ELSE LEFT(CAST(ROUND(T0.""TrsfrSum"",  " & vbCrLf
                sQuery &= " 					IFNULL((SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 				, LOCATE(T0.""TrsfrSum"", '.')  " & vbCrLf
                sQuery &= " 				+ IFNULL((SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "                	 END)  " & vbCrLf
                sQuery &= "    ELSE '' END)  " & vbCrLf

                'Credit Card Payment
                sQuery &= " || ''   " & vbCrLf

                'Cash Payment
                sQuery &= " || ( CASE  WHEN (T0.""CashAcct"" IS NOT NULL  " & vbCrLf
                sQuery &= " 		AND T0.""CashSum"" = ABS(T1.""Debit"" -  T1.""Credit"")) THEN  " & vbCrLf
                sQuery &= " 		'By Cash :  Amt '  " & vbCrLf
                sQuery &= " 		|| ( CASE  WHEN T0.""CashSumFC"" <> 0 THEN T0.""DocCurr""  " & vbCrLf
                sQuery &= "                     ELSE IFNULL( (SELECT T9.""MainCurncy""  FROM """ & oCompany.CompanyDB & """.OADM T9), '')  " & vbCrLf
                sQuery &= "              END)  " & vbCrLf
                sQuery &= " 		|| ' '  " & vbCrLf
                sQuery &= " 		|| ( CASE  WHEN T0.""CashSumFC"" <> 0  " & vbCrLf
                sQuery &= " 			THEN LEFT(CAST(ROUND(T0.""CashSumFC"" , IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 			    , LOCATE(T0.""CashSumFC"", '.')  " & vbCrLf
                sQuery &= " 				+ IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "               ELSE LEFT(CAST(ROUND(T0.""CashSum"", IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 			   , LOCATE(T0.""CashSum"", '.')  " & vbCrLf
                sQuery &= " 				+ IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "              END)  " & vbCrLf
                sQuery &= "  	ELSE '' END)  " & vbCrLf
                sQuery &= "  ELSE '' END) AS ""PayMeanRemarks"",  " & vbCrLf
                sQuery &= " '' AS ""PayTo"", T0.""Comments"",  " & vbCrLf
                sQuery &= "         CAST(T0.""ObjType"" AS varchar(3)) || ' ' || CAST(T0.""DocNum"" AS varchar(10)) AS ""OF_Link""  " & vbCrLf
                sQuery &= "  FROM """ & oCompany.CompanyDB & """.ORCT T0 " & vbCrLf
                sQuery &= "         INNER JOIN """ & oCompany.CompanyDB & """.JDT1 T1 ON T0.""TransId"" = T1.""TransId""  " & vbCrLf
                sQuery &= "         INNER JOIN """ & oCompany.CompanyDB & """.OACT T2 ON T1.""Account"" = T2.""AcctCode""  " & vbCrLf
                sQuery &= "  WHERE T0.""DocDate"" >= '" & oForm.Items.Item("dtPostFrom").Specific.Value & "'   " & vbCrLf
                sQuery &= "     AND T0.""DocDate"" <= '" & oForm.Items.Item("dtPostTo").Specific.Value & "' " & vbCrLf
                sQuery &= "  " & vbCrLf

                If oForm.Items.Item("txtORCTFr").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" >= " & sSORCT & " " & vbCrLf
                End If

                If oForm.Items.Item("txtORCTTo").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" <= " & sEORCT & "  " & vbCrLf
                End If

                sQuery &= " UNION ALL " & vbCrLf

                'OUTGOING RECEIPTS (ORCT)
                sQuery &= " SELECT T0.""ObjType"", CAST(T0.""DocNum"" AS varchar(10)) AS ""DocNumStr"", T0.""DocNum"", T0.""DocDate"", T0.""CounterRef"",   " & vbCrLf
                sQuery &= "     T0.""CardCode"", T0.""Canceled"", IFNULL(T2.""Segment_0"", T2.""AcctCode"") || (  " & vbCrLf
                sQuery &= "     CASE  WHEN T2.""Segment_1"" IS NULL THEN '' ELSE '-' || T2.""Segment_1"" END)   " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_2"" IS NULL THEN ''  ELSE '-' || T2.""Segment_2""  END)   " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_3"" IS NULL THEN ''  ELSE '-' || T2.""Segment_3""  END)   " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_4"" IS NULL THEN ''  ELSE '-' || T2.""Segment_4""  END)   " & vbCrLf
                sQuery &= "  	|| ( CASE  WHEN T2.""Segment_5"" IS NULL THEN ''  ELSE '-' || T2.""Segment_5"" 	END) AS ""AcctCode"", " & vbCrLf
                sQuery &= " 	( CASE  WHEN T1.""ShortName"" = T1.""Account"" THEN ''  ELSE T1.""ShortName"" 	END) AS ""ShortName"",   " & vbCrLf
                sQuery &= " 	T2.""AcctName"",  " & vbCrLf
                sQuery &= " 	( CASE WHEN T1.""Account"" = T1.""ShortName"" THEN ''   " & vbCrLf
                sQuery &= " 	  ELSE IFNULL( (SELECT T9.""CardName"" 	FROM """ & oCompany.CompanyDB & """.OCRD T9 WHERE T9.""CardCode"" = T1.""ShortName""), '')   " & vbCrLf
                sQuery &= "       END) AS ""CardName"" " & vbCrLf
                sQuery &= " 	, T1.""Debit"", T1.""Credit"", T1.""FCCurrency"", T1.""FCDebit"", T1.""FCCredit"",   " & vbCrLf
                sQuery &= "      CASE  WHEN T2.""Finanse"" = 'Y' THEN 1  ELSE 2  END AS ""BankLine"",  " & vbCrLf

                'Cheque Payment
                sQuery &= " ( CASE WHEN T2.""Finanse"" = 'Y' THEN (  " & vbCrLf
                sQuery &= "  	CASE  WHEN (T0.""CheckSum"" <> 0  " & vbCrLf
                sQuery &= " 		AND T0.""CheckSum"" = ABS(T1.""Debit"" -  T1.""Credit"")) THEN   " & vbCrLf
                sQuery &= " 		'By Cheque : Due on '   " & vbCrLf
                sQuery &= " 		|| IFNULL(  " & vbCrLf
                sQuery &= "                 	(SELECT   " & vbCrLf
                sQuery &= "                 		--TOP 1   " & vbCrLf
                sQuery &= "                 		CAST(T8.""DueDate"" AS varchar(12)) || '  Cheque ' || CAST(T8.""CheckNum"" AS varchar(10))   " & vbCrLf
                sQuery &= "                      	|| '  Amt ' || T8.""Currency"" || ' '  " & vbCrLf
                sQuery &= "  				        || LEFT(CAST(ROUND(T8.""CheckSum"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 			   	                , LOCATE(T8.""CheckSum"", '.')   " & vbCrLf
                sQuery &= " 				                    + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))   " & vbCrLf
                sQuery &= "                 	FROM """ & oCompany.CompanyDB & """.VPM1 T8  " & vbCrLf
                sQuery &= "                 	WHERE T8.""DocNum"" = T0.""DocNum""   " & vbCrLf
                sQuery &= " 			        AND T8.""CheckSum"" = ( CASE  WHEN T0.""CheckSumFC"" <> 0 THEN T0.""CheckSumFC""  " & vbCrLf
                sQuery &= " 			                                ELSE T0.""CheckSum"" END) " & vbCrLf
                sQuery &= " 			), '')  " & vbCrLf
                sQuery &= "      ELSE '' " & vbCrLf
                sQuery &= "      END) " & vbCrLf

                'Transfer Payment
                sQuery &= "  || ( CASE  WHEN (T0.""TrsfrSum"" <> 0   " & vbCrLf
                sQuery &= "  		AND T0.""TrsfrSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN  " & vbCrLf
                sQuery &= "  		'By Bank Transfer : ' || CAST(T0.""TrsfrDate"" AS varchar(12)) || '  '  " & vbCrLf
                sQuery &= " 		|| IFNULL(T0. ""TrsfrRef"", '') || '  Amt '   " & vbCrLf
                sQuery &= "  		|| ( CASE  WHEN T0.""TrsfrSumFC"" <> 0 THEN T0.""DocCurr""   " & vbCrLf
                sQuery &= "  		     ELSE IFNULL( (SELECT T9.""MainCurncy"" FROM """ & oCompany.CompanyDB & """.OADM T9), '')  " & vbCrLf
                sQuery &= "               END)   " & vbCrLf
                sQuery &= " 		|| ' '   " & vbCrLf
                sQuery &= " 		|| ( CASE  WHEN T0.""TrsfrSumFC"" <> 0   " & vbCrLf
                sQuery &= "  			    THEN LEFT(CAST(ROUND(T0.""TrsfrSumFC"", IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 				        , LOCATE(T0.""TrsfrSumFC"", '.')   " & vbCrLf
                sQuery &= " 					        + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))   " & vbCrLf
                sQuery &= "             ELSE LEFT(CAST(ROUND(T0.""TrsfrSum"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= "     				    , LOCATE(T0.""TrsfrSum"", '.')   " & vbCrLf
                sQuery &= "  					        + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "              END)  " & vbCrLf
                sQuery &= "     ELSE ''END) " & vbCrLf

                'Credit Card Payment
                sQuery &= "   || ''  " & vbCrLf

                'Cash Payment
                sQuery &= " || ( CASE WHEN (T0.""CashAcct"" IS NOT NULL   " & vbCrLf
                sQuery &= "  	    AND T0.""CashSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN  " & vbCrLf
                sQuery &= " 	    'By Cash :  Amt '   " & vbCrLf
                sQuery &= " 	    || ( CASE  WHEN T0.""CashSumFC"" <> 0 THEN T0.""DocCurr""   " & vbCrLf
                sQuery &= "             ELSE IFNULL( (SELECT T9.""MainCurncy""  FROM """ & oCompany.CompanyDB & """.OADM T9), '')  " & vbCrLf
                sQuery &= "             END)   " & vbCrLf
                sQuery &= " 	|| ' '   " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""CashSumFC"" <> 0 THEN   " & vbCrLf
                sQuery &= "  		    LEFT(CAST(ROUND(T0.""CashSumFC"" " & vbCrLf
                sQuery &= " 			, IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)), LOCATE(T0.""CashSumFC"", '.')   " & vbCrLf
                sQuery &= "         				+ IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "           ELSE LEFT(CAST(ROUND(T0.""CashSum""  " & vbCrLf
                sQuery &= " 			, IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)), LOCATE(T0.""CashSum"", '.')   " & vbCrLf
                sQuery &= " 				+ IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))   " & vbCrLf
                sQuery &= "           END)   " & vbCrLf
                sQuery &= "     ELSE ''  " & vbCrLf
                sQuery &= "     END)  " & vbCrLf
                sQuery &= " ELSE ''    " & vbCrLf
                sQuery &= " END) AS ""PayMeanRemarks"",  " & vbCrLf
                sQuery &= "  'Pay To : ' || T0.""Address"" AS ""PayTo"", T0.""Comments"",   " & vbCrLf
                sQuery &= "  CAST(T0.""ObjType"" AS varchar(3)) || ' ' || CAST(T0.""DocNum"" AS varchar(10)) AS ""OF_Link""   " & vbCrLf
                sQuery &= "      FROM """ & oCompany.CompanyDB & """.OVPM T0  " & vbCrLf
                sQuery &= "         INNER JOIN """ & oCompany.CompanyDB & """.JDT1 T1 ON T0.""TransId"" = T1.""TransId""   " & vbCrLf
                sQuery &= "         INNER JOIN """ & oCompany.CompanyDB & """.OACT T2 ON T1.""Account"" = T2.""AcctCode""  " & vbCrLf
                sQuery &= " WHERE T0.""DocDate"" >= '" & oForm.Items.Item("dtPostFrom").Specific.Value & "'   " & vbCrLf
                sQuery &= "     AND T0.""DocDate"" <= '" & oForm.Items.Item("dtPostTo").Specific.Value & "'   " & vbCrLf

                If oForm.Items.Item("txtOVPMFr").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" >= " & sSOVPM & "  " & vbCrLf
                End If

                If oForm.Items.Item("txtOVPMTo").Specific.Value.ToString.Length > 0 Then
                    sQuery &= "  AND T0.""DocNum"" <= " & sEOVPM & "   " & vbCrLf
                End If

                ' --- additional section for cancellation
                sQuery &= " UNION ALL " & vbCrLf

                'INCOMING PAYMENT CANCELLATION (ORCT)
                sQuery &= " SELECT T0.""ObjType"", 'C' || CAST(T0.""DocNum"" AS varchar(10)) AS ""DocNumStr"", T0.""DocNum"", T0.""DocDate"",  " & vbCrLf
                sQuery &= "         T0.""CounterRef"", T0.""CardCode"", T0.""Canceled"", IFNULL(T2.""Segment_0"", T2.""AcctCode"") || ( " & vbCrLf
                sQuery &= "         CASE WHEN T2.""Segment_1"" IS NULL THEN '' ELSE '-' || T2.""Segment_1""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T2.""Segment_2"" IS NULL THEN ''  ELSE '-' || T2.""Segment_2""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T2.""Segment_3"" IS NULL THEN ''  ELSE '-' || T2.""Segment_3"" END)  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T2.""Segment_4"" IS NULL THEN ''  ELSE '-' || T2.""Segment_4""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T2.""Segment_5"" IS NULL THEN ''  ELSE '-' || T2.""Segment_5""  END) AS ""AcctCode"" " & vbCrLf
                sQuery &= " 	, ( CASE  WHEN T1.""ShortName"" = T1.""Account"" THEN '' ELSE T1.""ShortName""  END) AS ""ShortName"" " & vbCrLf
                sQuery &= " 	, T2.""AcctName"" " & vbCrLf
                sQuery &= " 	, ( CASE  WHEN T1.""Account"" = T1.""ShortName"" THEN ''  " & vbCrLf
                sQuery &= "         ELSE IFNULL( (SELECT T9.""CardName""  FROM """ & oCompany.CompanyDB & """.OCRD T9  WHERE T9.""CardCode"" = T1.""ShortName""), '') " & vbCrLf
                sQuery &= "         END) AS ""CardName"",  " & vbCrLf
                sQuery &= " 	T1.""Debit"", T1.""Credit"", T1.""FCCurrency"", T1.""FCDebit"", T1.""FCCredit"",  " & vbCrLf
                sQuery &= "         CASE WHEN T2.""Finanse"" = 'Y' THEN 1  ELSE 2 END AS ""BankLine"",  " & vbCrLf

                'Cheque Payment
                sQuery &= " ( CASE WHEN T2.""Finanse"" = 'Y' THEN (  " & vbCrLf
                sQuery &= "  	CASE WHEN (T0.""CheckSum"" <> 0  " & vbCrLf
                sQuery &= " 		    AND T0.""CheckSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN   " & vbCrLf
                sQuery &= "  		'By Cheque : Due on ' " & vbCrLf
                sQuery &= "  		|| IFNULL( " & vbCrLf
                sQuery &= "                 	(SELECT   " & vbCrLf
                sQuery &= "                  	--TOP 1  " & vbCrLf
                sQuery &= "  			        CAST(T8.""DueDate"" AS varchar(12)) || '  Cheque ' || CAST(T8.""CheckNum"" AS varchar(10))  " & vbCrLf
                sQuery &= "                     	|| '  Amt ' || T8.""Currency"" || ' '   " & vbCrLf
                sQuery &= "  			        || LEFT(CAST(ROUND(T8.""CheckSum"", IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 				                    , LOCATE(T8.""CheckSum"", '.')    " & vbCrLf
                sQuery &= "  					                + IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "                 	FROM """ & oCompany.CompanyDB & """.RCT1 T8   " & vbCrLf
                sQuery &= "                 	WHERE T8.""DocNum"" = T0.""DocNum""   " & vbCrLf
                sQuery &= " 				        AND T8.""CheckSum"" = ( CASE WHEN T0.""CheckSumFC"" <> 0 THEN T0.""CheckSumFC""   " & vbCrLf
                sQuery &= "                                               ELSE T0.""CheckSum"" END)) " & vbCrLf
                sQuery &= "  		  , '') " & vbCrLf
                sQuery &= "       ELSE ''  END)  " & vbCrLf

                'Transfer Payment
                sQuery &= " || ( CASE WHEN (T0.""TrsfrSum"" <> 0   " & vbCrLf
                sQuery &= "  		AND T0.""TrsfrSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN  " & vbCrLf
                sQuery &= " 	'By Bank Transfer : ' || CAST(T0.""TrsfrDate"" AS varchar(12)) || '  '   " & vbCrLf
                sQuery &= " 	|| IFNULL(T0.""TrsfrRef"", '') || '  Amt '   " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""TrsfrSumFC"" <> 0 THEN T0.""DocCurr""   " & vbCrLf
                sQuery &= "    	     ELSE IFNULL( (SELECT T9.""MainCurncy""  FROM """ & oCompany.CompanyDB & """.OADM T9), '')   " & vbCrLf
                sQuery &= "          END)   " & vbCrLf
                sQuery &= " 	|| ' '   " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""TrsfrSumFC"" <> 0 THEN   " & vbCrLf
                sQuery &= "  		   LEFT(CAST(ROUND(T0.""TrsfrSumFC"" , IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= "  		  , LOCATE(T0.""TrsfrSumFC"", '.')  " & vbCrLf
                sQuery &= " 			    + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))   " & vbCrLf
                sQuery &= " 		ELSE LEFT(CAST(ROUND(T0.""TrsfrSum"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= " 		      , LOCATE(T0.""TrsfrSum"", '.')   " & vbCrLf
                sQuery &= " 			    + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))   " & vbCrLf
                sQuery &= "          END)  " & vbCrLf
                sQuery &= "     ELSE '' END)  " & vbCrLf

                'Credit Card Payment
                sQuery &= " || ''  " & vbCrLf

                'Cash Payment
                sQuery &= " || ( CASE WHEN (T0.""CashAcct"" IS NOT NULL AND T0.""CashSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN   " & vbCrLf
                sQuery &= "  	'By Cash :  Amt '  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""CashSumFC"" <> 0 THEN T0.""DocCurr""   " & vbCrLf
                sQuery &= " 		 ELSE IFNULL( (SELECT T9.""MainCurncy"" FROM """ & oCompany.CompanyDB & """.OADM T9), '')   " & vbCrLf
                sQuery &= "          END)   " & vbCrLf
                sQuery &= " 	|| ' '  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""CashSumFC"" <> 0 THEN   " & vbCrLf
                sQuery &= " 		    LEFT(CAST(ROUND(T0.""CashSumFC"", IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= "  			, LOCATE(T0.""CashSumFC"", '.')  " & vbCrLf
                sQuery &= "  				+ IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= " 	    ELSE LEFT(CAST(ROUND(T0.""CashSum"", IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= " 			, LOCATE(T0.""CashSum"", '.')   " & vbCrLf
                sQuery &= " 				+ IFNULL( (SELECT T9.""SumDec"" FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) 		  " & vbCrLf
                sQuery &= "  	    END) " & vbCrLf
                sQuery &= "      ELSE ''END)  " & vbCrLf
                sQuery &= " ELSE '' END) AS ""PayMeanRemarks""  " & vbCrLf
                sQuery &= " , '' AS ""PayTo"", T0.""Comments"", CAST(T0.""ObjType"" AS varchar(3)) || ' ' || CAST(T0.""DocNum"" AS varchar(10)) AS ""OF_Link""   " & vbCrLf
                sQuery &= "  FROM """ & oCompany.CompanyDB & """.ORCT T0   " & vbCrLf
                sQuery &= "  INNER JOIN """ & oCompany.CompanyDB & """.OJDT J ON T0.""TransId"" = J.""StornoToTr""  " & vbCrLf
                sQuery &= "  INNER JOIN """ & oCompany.CompanyDB & """.JDT1 T1 ON J.""TransId"" = T1.""TransId""  " & vbCrLf
                sQuery &= "  INNER JOIN """ & oCompany.CompanyDB & """.OACT T2 ON T1.""Account"" = T2.""AcctCode""  " & vbCrLf
                sQuery &= "  WHERE T0.""DocDate"" >= '" & oForm.Items.Item("dtPostFrom").Specific.Value & "'   " & vbCrLf
                sQuery &= "     AND T0.""DocDate"" <= '" & oForm.Items.Item("dtPostTo").Specific.Value & "'  " & vbCrLf
                sQuery &= " 	AND J.""StornoToTr"" IS NOT NULL  " & vbCrLf
                sQuery &= " 	AND T0.""Canceled"" = 'Y'  " & vbCrLf

                If oForm.Items.Item("txtORCTFr").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" >= " & sSORCT & "   " & vbCrLf
                End If

                If oForm.Items.Item("txtORCTTo").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" <= " & sEORCT & "   " & vbCrLf
                End If

                sQuery &= " UNION ALL " & vbCrLf

                'OUTGOING PAYMENTS CCANCELLATION (OVPM)

                sQuery &= " SELECT T0.""ObjType"", 'C' || CAST(T0.""DocNum"" AS varchar(10)) AS ""DocNumStr"", T0.""DocNum"", T0.""DocDate"",  " & vbCrLf
                sQuery &= "         T0.""CounterRef"", T0.""CardCode"", T0.""Canceled"", IFNULL(T2.""Segment_0"", T2.""AcctCode"") || ( " & vbCrLf
                sQuery &= "         CASE  WHEN T2.""Segment_1"" IS NULL THEN ''  ELSE '-' || T2.""Segment_1""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_2"" IS NULL THEN ''  ELSE '-' || T2.""Segment_2""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_3"" IS NULL THEN ''  ELSE '-' || T2.""Segment_3""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_4"" IS NULL THEN ''  ELSE '-' || T2.""Segment_4""  END)  " & vbCrLf
                sQuery &= " 	|| ( CASE  WHEN T2.""Segment_5"" IS NULL THEN ''  ELSE '-' || T2.""Segment_5""  END) AS ""AcctCode"" " & vbCrLf
                sQuery &= " 	, ( CASE  WHEN T1.""ShortName"" = T1.""Account"" THEN ''  ELSE T1.""ShortName""   END) AS ""ShortName"" " & vbCrLf
                sQuery &= " 	, T2.""AcctName"" " & vbCrLf
                sQuery &= " 	, ( CASE  WHEN T1.""Account"" = T1.""ShortName"" THEN ''  " & vbCrLf
                sQuery &= "          ELSE IFNULL( (SELECT T9.""CardName""  FROM """ & oCompany.CompanyDB & """.OCRD T9  WHERE T9.""CardCode"" = T1.""ShortName""), '')  " & vbCrLf
                sQuery &= "          END) AS ""CardName"" " & vbCrLf
                sQuery &= " 	, T1.""Debit"", T1.""Credit"", T1.""FCCurrency"", T1.""FCDebit"", T1.""FCCredit"",  " & vbCrLf
                sQuery &= "         CASE WHEN T2.""Finanse"" = 'Y' THEN 1  ELSE 2  END AS ""BankLine"" " & vbCrLf

                'Cheque Payment
                sQuery &= " , ( CASE WHEN T2.""Finanse"" = 'Y' THEN ( " & vbCrLf
                sQuery &= " 	    CASE WHEN (T0.""CheckSum"" <> 0 AND T0.""CheckSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN  " & vbCrLf
                sQuery &= " 		'By Cheque : Due on '  " & vbCrLf
                sQuery &= " 		|| IFNULL( " & vbCrLf
                sQuery &= "         (SELECT  " & vbCrLf
                sQuery &= "             --TOP 1  " & vbCrLf
                sQuery &= " 			CAST(T8.""DueDate"" AS varchar(12)) || '  Cheque ' || CAST(T8.""CheckNum"" AS varchar(10))  " & vbCrLf
                sQuery &= "             || '  Amt ' || T8.""Currency"" || ' '  " & vbCrLf
                sQuery &= " 			|| LEFT(CAST(ROUND(T8.""CheckSum"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20)) " & vbCrLf
                sQuery &= " 				, LOCATE(T8.""CheckSum"", '.')  " & vbCrLf
                sQuery &= " 					+ IFNULL((SELECT T9.""SumDec"" 	FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "           FROM """ & oCompany.CompanyDB & """.VPM1 T8  " & vbCrLf
                sQuery &= "           WHERE T8.""DocNum"" = T0.""DocNum""  " & vbCrLf
                sQuery &= " 			AND T8.""CheckSum"" = ( CASE  WHEN T0.""CheckSumFC"" <> 0 THEN T0.""CheckSumFC""  " & vbCrLf
                sQuery &= "                         			ELSE T0.""CheckSum""  " & vbCrLf
                sQuery &= "                     				END)) " & vbCrLf
                sQuery &= " 		, '')  " & vbCrLf
                sQuery &= "          ELSE '' END)  " & vbCrLf

                'Transfer Payment
                sQuery &= " || ( CASE  WHEN (T0.""TrsfrSum"" <> 0 AND T0.""TrsfrSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN   " & vbCrLf
                sQuery &= "  	'By Bank Transfer : ' || CAST(T0.""TrsfrDate"" AS varchar(12)) || '  '  " & vbCrLf
                sQuery &= "  	|| IFNULL(T0.""TrsfrRef"", '') || '  Amt '  " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""TrsfrSumFC"" <> 0 THEN T0.""DocCurr""   " & vbCrLf
                sQuery &= " 		ELSE IFNULL( (SELECT T9.""MainCurncy""  FROM """ & oCompany.CompanyDB & """.OADM T9), '')   " & vbCrLf
                sQuery &= "         END)  " & vbCrLf
                sQuery &= " 	|| ' '   " & vbCrLf
                sQuery &= " 	|| ( CASE WHEN T0.""TrsfrSumFC"" <> 0   " & vbCrLf
                sQuery &= " 		 THEN LEFT(CAST(ROUND(T0.""TrsfrSumFC"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= " 			, LOCATE(T0.""TrsfrSumFC"", '.')   " & vbCrLf
                sQuery &= " 				+ IFNULL( (SELECT T9.""SumDec""   FROM """ & oCompany.CompanyDB & """.OADM T9), 2))   " & vbCrLf
                sQuery &= " 	     ELSE LEFT(CAST(ROUND(T0.""TrsfrSum"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= "  			    , LOCATE(T0.""TrsfrSum"", '.')  " & vbCrLf
                sQuery &= "  				    + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "          END)  " & vbCrLf
                sQuery &= "    ELSE ''  " & vbCrLf
                sQuery &= "    END)   " & vbCrLf

                'Credit Card Payment
                sQuery &= " || ''   " & vbCrLf

                'Cash Payment
                sQuery &= " || ( CASE WHEN (T0.""CashAcct"" IS NOT NULL AND T0.""CashSum"" = ABS(T1.""Debit"" - T1.""Credit"")) THEN   " & vbCrLf
                sQuery &= " 		'By Cash :  Amt '   " & vbCrLf
                sQuery &= " 		|| ( CASE  WHEN T0.""CashSumFC"" <> 0 THEN T0.""DocCurr""   " & vbCrLf
                sQuery &= "               ELSE IFNULL( (SELECT T9.""MainCurncy""  FROM """ & oCompany.CompanyDB & """.OADM T9), '')  " & vbCrLf
                sQuery &= "               END)  " & vbCrLf
                sQuery &= "  		|| ' '  " & vbCrLf
                sQuery &= "  		|| ( CASE WHEN T0.""CashSumFC"" <> 0 THEN  " & vbCrLf
                sQuery &= " 			    LEFT(CAST(ROUND(T0.""CashSumFC"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= " 				    , LOCATE(T0.""CashSumFC"", '.')   " & vbCrLf
                sQuery &= "  					    + IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "       ELSE LEFT(CAST(ROUND(T0.""CashSum"", IFNULL( (SELECT T9.""SumDec""  FROM """ & oCompany.CompanyDB & """.OADM T9), 2)) AS varchar(20))  " & vbCrLf
                sQuery &= "  				, LOCATE(T0.""CashSum"", '.')  " & vbCrLf
                sQuery &= "  					+ IFNULL( (SELECT T9.""SumDec""   FROM """ & oCompany.CompanyDB & """.OADM T9), 2))  " & vbCrLf
                sQuery &= "        END)  " & vbCrLf
                sQuery &= " 	ELSE ''   " & vbCrLf
                sQuery &= " 	END)   " & vbCrLf
                sQuery &= " ELSE ''   " & vbCrLf
                sQuery &= " END) AS ""PayMeanRemarks""  " & vbCrLf
                sQuery &= " , 'Pay To : ' || T0.""Address"" AS ""PayTo"", T0.""Comments"",   " & vbCrLf
                sQuery &= "  CAST(T0.""ObjType"" AS varchar(3)) || ' ' || CAST(T0.""DocNum"" AS varchar(10)) AS ""OF_Link""  " & vbCrLf
                sQuery &= "  FROM """ & oCompany.CompanyDB & """.OVPM T0  " & vbCrLf
                sQuery &= "          INNER JOIN """ & oCompany.CompanyDB & """.OJDT J ON T0.""TransId"" = J.""StornoToTr""  " & vbCrLf
                sQuery &= "          INNER JOIN """ & oCompany.CompanyDB & """.JDT1 T1 ON J.""TransId"" = T1.""TransId""  " & vbCrLf
                sQuery &= "          INNER JOIN """ & oCompany.CompanyDB & """.OACT T2 ON T1.""Account"" = T2.""AcctCode""  " & vbCrLf
                sQuery &= "  WHERE T0.""DocDate"" >= '" & oForm.Items.Item("dtPostFrom").Specific.Value & "'  " & vbCrLf
                sQuery &= "  	 AND T0.""DocDate"" <= '" & oForm.Items.Item("dtPostTo").Specific.Value & "'  " & vbCrLf
                sQuery &= "      AND J.""StornoToTr"" IS NOT NULL  " & vbCrLf
                sQuery &= "  	 AND T0.""Canceled"" = 'Y'  " & vbCrLf

                If oForm.Items.Item("txtOVPMFr").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" >= " & sSOVPM & "   " & vbCrLf
                End If

                If oForm.Items.Item("txtOVPMTo").Specific.Value.ToString.Length > 0 Then
                    sQuery &= " AND T0.""DocNum"" <= " & sEOVPM & "  " & vbCrLf
                End If
                ' --- end - additional section for cancellation

                sQuery &= "  ) AS X  " & vbCrLf

                Select Case oForm.DataSources.UserDataSources.Item("cbType").ValueEx
                    Case "-"
                        '' do nothing
                    Case "24"
                        sQuery &= " WHERE X.""ObjType"" = '24' "
                    Case "46"
                        sQuery &= " WHERE X.""ObjType"" = '46' "
                End Select

                sQuery &= " ORDER BY X.""DocDate"", X.""ObjType"" DESC, CAST(X.""DocNum"" AS decimal(15, 5)) "
            End With

            Dim ProviderName As String = "System.Data.Odbc"
            Dim _DbProviderFactoryObject As DbProviderFactory

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dt)


            ' HANA
            SBO_Application.StatusBar.SetText("Collecting Details...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSelect &= " SELECT T1.""ObjType"", T0.""DocNum"",  "
            sSelect &= "     CASE T0.""InvType""  "
            sSelect &= "         WHEN 13 THEN 'IN'  "
            sSelect &= "         WHEN 203 THEN 'DPI'  "
            sSelect &= "         WHEN 14 THEN 'CN'  "
            sSelect &= "         WHEN 30 THEN 'JE'  "
            sSelect &= "         ELSE '??'  "
            sSelect &= "     END AS ""OF_ObjType"",  "
            sSelect &= "     CASE T0.""InvType""  "
            sSelect &= "         WHEN 13 THEN IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OINV T9  "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 203 THEN IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10)) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ODPI T9  "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 14 THEN IFNULL("
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ORIN T9  "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 30 THEN IFNULL("
            sSelect &= "         (SELECT CAST(T9.""TransId"" AS varchar(10)) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OJDT T9 "
            sSelect &= "         WHERE T9.""TransId"" = T0.""DocEntry""), '') "
            sSelect &= "         ELSE '??' "
            sSelect &= "     END AS ""OF_DocNum"", "
            sSelect &= "     CASE T0.""InvType""  "
            sSelect &= "         WHEN 13 THEN 'IN ' || IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OINV T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "         WHEN 203 THEN 'DPI ' || IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ODPI T9  "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "         WHEN 14 THEN 'CN ' || IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10)) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ORIN T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 30 THEN 'JE ' || IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""TransId"" AS varchar(10)) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OJDT T9 "
            sSelect &= "         WHERE T9.""TransId"" = T0.""DocEntry""), '') "
            sSelect &= "         ELSE CAST(T0.""InvType"" AS varchar(4)) || ' ' || CAST(T0.""DocEntry"" AS varchar(10)) "
            sSelect &= "     END AS ""OF_DocNumStr"", T0.""DocEntry"" AS ""OF_DocEntry"", "
            sSelect &= "     CASE T0.""InvType"" "
            sSelect &= "         WHEN 13 THEN IFNULL( "
            sSelect &= "         (SELECT T9.""DocCur"" "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OINV T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 203 THEN IFNULL( "
            sSelect &= "         (SELECT T9.""DocCur""  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ODPI T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 14 THEN IFNULL( "
            sSelect &= "         (SELECT T9.""DocCur""  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ORIN T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '') "
            sSelect &= "         WHEN 30 THEN IFNULL( "
            sSelect &= "         (SELECT T9.""FCCurrency"" "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".JDT1 T9 "
            sSelect &= "         WHERE T9.""TransId"" = T0.""DocEntry"" AND T9.""ShortName"" = T1.""CardCode"" AND T9.""Line_ID"" = T0.""DocLine""), '') "
            sSelect &= "         ELSE '??' "
            sSelect &= "     END AS ""OF_DocCurr"", "
            sSelect &= "     CASE T0.""InvType"" "
            sSelect &= "         WHEN 13 THEN IFNULL( "
            sSelect &= "         (SELECT ( "
            sSelect &= "             CASE "
            sSelect &= "                 WHEN T9.""DocTotalFC"" = 0 THEN T9.""DocTotal"" "
            sSelect &= "                 ELSE T9.""DocTotalFC"" "
            sSelect &= "             END) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OINV T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), 0) "
            sSelect &= "         WHEN 203 THEN IFNULL( "
            sSelect &= "         (SELECT ( "
            sSelect &= "             CASE  "
            sSelect &= "                 WHEN T9.""DocTotalFC"" = 0 THEN T9.""DocTotal""  "
            sSelect &= "                 ELSE T9.""DocTotalFC""  "
            sSelect &= "             END) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ODPI T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), 0) "
            sSelect &= "         WHEN 14 THEN IFNULL( "
            sSelect &= "         (SELECT ( "
            sSelect &= "             CASE "
            sSelect &= "                 WHEN T9.""DocTotalFC"" = 0 THEN T9.""DocTotal"" "
            sSelect &= "                 ELSE T9.""DocTotalFC"" "
            sSelect &= "             END) "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ORIN T9 "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), 0) "
            sSelect &= "         WHEN 30 THEN IFNULL( "
            sSelect &= "         (SELECT ( "
            sSelect &= "             CASE "
            sSelect &= "                 WHEN (T9.""FCDebit"" - T9.""FCCredit"") = 0 THEN (T9.""Debit"" - T9.""Credit"") "
            sSelect &= "                 ELSE (T9.""FCDebit"" - T9.""FCCredit"")   "
            sSelect &= "             END)   "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".JDT1 T9   "
            sSelect &= "          WHERE T9.""TransId"" = T0.""DocEntry"" AND IFNULL(T9.""ShortName"", '') = IFNULL(T1.""CardCode"", '') AND "
            sSelect &= "               T9.""Line_ID"" = T0.""DocLine""), 0)  "
            sSelect &= "          ELSE 0 "
            sSelect &= "      END AS ""OF_DocAmt"",  "
            sSelect &= "      CASE T0.""InvType""  "
            sSelect &= "          WHEN 14 THEN T0.""AppliedFC"" * -1  "
            sSelect &= "          ELSE T0.""AppliedFC""  "
            sSelect &= "      END AS ""OF_FCAmt"",  "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "          WHEN 14 THEN T0.""SumApplied"" * -1  "
            sSelect &= "         ELSE T0.""SumApplied""   "
            sSelect &= "      END AS ""OF_LCAmt"", CAST(T0.""ObjType"" AS varchar(3)) || ' ' || CAST(T0.""DocNum"" AS varchar(10)) AS ""OF_Link""  "
            sSelect &= "  FROM " & oCompany.CompanyDB & ".RCT2 T0 "
            sSelect &= "  LEFT OUTER JOIN " & oCompany.CompanyDB & ".ORCT T1 "
            sSelect &= "  ON T0.""DocNum"" = T1.""DocNum""  "
            sSelect &= " WHERE T1.""DocDate"" >= '" & oForm.DataSources.UserDataSources.Item("dtPostFrom").ValueEx & "' "
            sSelect &= " AND T1.""DocDate"" <= '" & oForm.DataSources.UserDataSources.Item("dtPostTo").ValueEx & "' "
            sSelect &= " AND T1.""DocNum"" >= '" & sSORCT & "' "
            sSelect &= " AND T1.""DocNum"" <= '" & sEORCT & "' "

            sSelect &= " UNION ALL"

            sSelect &= " SELECT T1.""ObjType"", T0.""DocNum"",   "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "         WHEN 18 THEN 'PU'   "
            sSelect &= "         WHEN 204 THEN 'DPO'   "
            sSelect &= "         WHEN 19 THEN 'PC'  "
            sSelect &= "         WHEN 30 THEN 'JE'  "
            sSelect &= "         ELSE '??'   "
            sSelect &= "     END AS ""OF_ObjType"",   "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "          WHEN 18 THEN IFNULL( "
            sSelect &= "          (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OPCH T9   "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "         WHEN 204 THEN IFNULL(  "
            sSelect &= "         (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ODPO T9   "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '')   "
            sSelect &= "         WHEN 19 THEN IFNULL(  "
            sSelect &= "          (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".ORPC T9  "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "         WHEN 30 THEN IFNULL(  "
            sSelect &= "          (SELECT CAST(T9.""TransId"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OJDT T9   "
            sSelect &= "          WHERE T9.""TransId"" = T0.""DocEntry""), '')  "
            sSelect &= "         ELSE '??'   "
            sSelect &= "      END AS ""OF_DocNum"",  "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "         WHEN 18 THEN 'PU ' || IFNULL(  "
            sSelect &= "          (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".OPCH T9  "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "          WHEN 204 THEN 'DPO ' || IFNULL( "
            sSelect &= "          (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".ODPO T9  "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '')   "
            sSelect &= "          WHEN 19 THEN 'PC ' || IFNULL( "
            sSelect &= "          (SELECT CAST(T9.""DocNum"" AS varchar(10))  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ORPC T9   "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "          WHEN 30 THEN 'JE ' || IFNULL( "
            sSelect &= "         (SELECT CAST(T9.""TransId"" AS varchar(10))   "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".OJDT T9   "
            sSelect &= "         WHERE T9.""TransId"" = T0.""DocEntry""), '')   "
            sSelect &= "          ELSE CAST(T0.""InvType"" AS varchar(4)) || ' ' || CAST(T0.""DocEntry"" AS varchar(10))  "
            sSelect &= "     END AS ""OF_DocNumStr"", T0.""DocEntry"" AS ""OF_DocEntry"",   "
            sSelect &= "      CASE T0.""InvType""  "
            sSelect &= "          WHEN 18 THEN IFNULL( "
            sSelect &= "          (SELECT T9.""DocCur""  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".OPCH T9  "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "          WHEN 204 THEN IFNULL( "
            sSelect &= "          (SELECT T9.""DocCur""  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".ODPO T9  "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), '')  "
            sSelect &= "         WHEN 19 THEN IFNULL(  "
            sSelect &= "         (SELECT T9.""DocCur""   "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ORPC T9   "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), '')   "
            sSelect &= "          WHEN 30 THEN IFNULL( "
            sSelect &= "          (SELECT T9.""FCCurrency""  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".JDT1 T9  "
            sSelect &= "          WHERE T9.""TransId"" = T0.""DocEntry"" AND T9.""ShortName"" = T1.""CardCode"" AND T9.""Line_ID"" = T0.""DocLine""), '')  "
            sSelect &= "         ELSE '??'   "
            sSelect &= "     END AS ""OF_DocCurr"",   "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "          WHEN 18 THEN IFNULL( "
            sSelect &= "          (SELECT ( "
            sSelect &= "              CASE  "
            sSelect &= "                  WHEN T9.""DocTotalFC"" = 0 THEN T9.""DocTotal""  "
            sSelect &= "                  ELSE T9.""DocTotalFC""  "
            sSelect &= "              END)  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".OPCH T9  "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), 0)  "
            sSelect &= "          WHEN 204 THEN IFNULL( "
            sSelect &= "         (SELECT (  "
            sSelect &= "             CASE   "
            sSelect &= "                  WHEN T9.""DocTotalFC"" = 0 THEN T9.""DocTotal""  "
            sSelect &= "                  ELSE T9.""DocTotalFC""   "
            sSelect &= "              END)  "
            sSelect &= "         FROM " & oCompany.CompanyDB & ".ODPO T9   "
            sSelect &= "          WHERE T9.""DocEntry"" = T0.""DocEntry""), 0)  "
            sSelect &= "          WHEN 19 THEN IFNULL( "
            sSelect &= "         (SELECT (  "
            sSelect &= "             CASE   "
            sSelect &= "                  WHEN T9.""DocTotalFC"" = 0 THEN T9.""DocTotal""  "
            sSelect &= "                  ELSE T9.""DocTotalFC""   "
            sSelect &= "              END)  "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".ORPC T9  "
            sSelect &= "         WHERE T9.""DocEntry"" = T0.""DocEntry""), 0)   "
            sSelect &= "          WHEN 30 THEN IFNULL( "
            sSelect &= "         (SELECT (  "
            sSelect &= "             CASE   "
            sSelect &= "                 WHEN (T9.""FCDebit"" - T9.""FCCredit"") = 0 THEN (T9.""Debit"" - T9.""Credit"") "
            sSelect &= "                  ELSE (T9.""FCDebit"" - T9.""FCCredit"")  "
            sSelect &= "             END)   "
            sSelect &= "          FROM " & oCompany.CompanyDB & ".JDT1 T9  "
            sSelect &= "         WHERE T9.""TransId"" = T0.""DocEntry"" AND IFNULL(T9.""ShortName"", '') = IFNULL(T1.""CardCode"", '') AND  "
            sSelect &= "               T9.""Line_ID"" = T0.""DocLine""), 0)  "
            sSelect &= "         ELSE 0   "
            sSelect &= "     END AS ""OF_DocAmt"",   "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "         WHEN 19 THEN T0.""AppliedFC"" * -1   "
            sSelect &= "         ELSE T0.""AppliedFC""   "
            sSelect &= "      END AS ""OF_FCAmt"",  "
            sSelect &= "     CASE T0.""InvType""   "
            sSelect &= "         WHEN 19 THEN T0.""SumApplied"" * -1   "
            sSelect &= "         ELSE T0.""SumApplied""   "
            sSelect &= "     END AS ""OF_LCAmt"", CAST(T0.""ObjType"" AS varchar(3)) || ' ' || CAST(T0.""DocNum"" AS varchar(10)) AS ""OF_Link""   "
            sSelect &= " FROM " & oCompany.CompanyDB & ".VPM2 T0 LEFT OUTER JOIN   "
            sSelect &= "     " & oCompany.CompanyDB & ".OVPM T1 ON T0.""DocNum"" = T1.""DocNum""  "
            sSelect &= " WHERE T1.""DocDate"" >= '" & oForm.DataSources.UserDataSources.Item("dtPostFrom").ValueEx & "' "
            sSelect &= " AND T1.""DocDate"" <= '" & oForm.DataSources.UserDataSources.Item("dtPostTo").ValueEx & "' "
            sSelect &= " AND T1.""DocNum"" >= '" & sSOVPM & "' "
            sSelect &= " AND T1.""DocNum"" <= '" & sEOVPM & "' "

            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sSelect
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dt_DETAILS)
            dbConn.Close()

            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            If dbConn.State = ConnectionState.Open Then
                dbConn.Close()
            End If
            SBO_Application.MessageBox("[ExecuteProcedure]:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.RecpPaym)
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
    Private Function ValidateParameter() As Boolean
        Try
            oForm.ActiveItem = "txtORCTFr"
            Dim sStart As String = ""
            Dim sEnd As String = ""

            sStart = oForm.DataSources.UserDataSources.Item("dtPostFrom").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("dtPostTo").ValueEx

            If (sStart.Length = 0) Then
                SBO_Application.StatusBar.SetText("Please enter Starting Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "dtPostFrom"
                Return False
            End If
            If (sEnd.Length = 0) Then
                SBO_Application.StatusBar.SetText("Please enter Ending Posting Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "dtPostTo"
                Return False
            End If

            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Date From is greater than Date To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "dtPostTo"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtORCTFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtORCTTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Receipt Doc Num From is greater than Doc Num To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtORCTFr"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtOVPMFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtOVPMTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Payment Doc Num From is greater than Doc Num To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtOVPMFr"
                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[frmRecpPaym].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        Select Case pVal.ItemUID
                            Case "txtOVPMFr", "txtOVPMTo"
                                oEdit = oForm.Items.Item(pVal.ItemUID).Specific
                                If (oEdit.String.ToString.Trim = "") And (pVal.CharPressed = 9) Then
                                    SBO_Application.SendKeys("+{F2}")
                                    Return False
                                End If
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "btnPrint" Then
                            If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                Return ValidateParameter()
                            End If
                        End If
                End Select

            Else
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then
                    Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvent = pVal
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oDataTable = oCFLEvent.SelectedObjects
                    If (Not oDataTable Is Nothing) Then
                        Dim sTemp As String = String.Empty
                        Select Case oCFLEvent.ChooseFromListUID
                            Case "cflSORCT"
                                sTemp = oDataTable.GetValue("DocNum", 0)
                                oForm.DataSources.UserDataSources.Item("txtORCTFr").ValueEx = sTemp
                                Exit Select
                            Case "cflEORCT"
                                sTemp = oDataTable.GetValue("DocNum", 0)
                                oForm.DataSources.UserDataSources.Item("txtORCTTo").ValueEx = sTemp
                                Exit Select
                            Case "cflSOVPM"
                                sTemp = oDataTable.GetValue("DocNum", 0)
                                oForm.DataSources.UserDataSources.Item("txtOVPMFr").ValueEx = sTemp
                                Exit Select
                            Case "cflEOVPM"
                                sTemp = oDataTable.GetValue("DocNum", 0)
                                oForm.DataSources.UserDataSources.Item("txtOVPMTo").ValueEx = sTemp
                                Exit Select
                            Case Else
                                Exit Select
                        End Select
                        Return True
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.ItemUID = "btnPrint" Then
                        If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                            myThread.Start()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[frmRecpPaym].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class