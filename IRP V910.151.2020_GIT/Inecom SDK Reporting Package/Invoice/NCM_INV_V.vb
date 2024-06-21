Imports System.IO
Imports System.IO.Compression
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Globalization
Imports System.Xml
Imports System.Data.Common

Public Class NCM_INV_V

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oMtrx As SAPbouiCOM.Matrix
    Private g_bINV_SharedFile As Boolean = False
    Private g_sINV_SharedFile As String = ""
    Private g_sSelectedDocType As String = ""
    Private myThread As System.Threading.Thread
    Private g_iSelectedLine As Integer = 0
    Private g_sColUID As String = ""
    Private g_sLastFolder As String = ""
    Private g_sZippedFile As String = ""
    Private g_bSelectAll As String = ""

    Private dtEInvoice As System.Data.DataTable
    Private dtOADM As System.Data.DataTable

    Private dtADM1 As System.Data.DataTable
    Private dtOCRD As System.Data.DataTable
    Private dtOCTG As System.Data.DataTable
    Private dtOSLP As System.Data.DataTable
    Private dtOPRJ As System.Data.DataTable
    Private dtOEXD As System.Data.DataTable
    Private dtSRADDRESS3 As System.Data.DataTable
    Private dtSRREASON As System.Data.DataTable

    Private dtOINV As System.Data.DataTable
    Private dtINV1 As System.Data.DataTable
    Private dtINV3 As System.Data.DataTable
    Private dtINV10 As System.Data.DataTable

    Private g_sOINV_Query As String = ""
    Private g_sINV1_Query As String = ""
    Private g_sINV3_Query As String = ""
    Private g_sINV10_Query As String = ""
    Private g_sOINV_Last_Query As String = ""
    Private g_sCOMPANYSETTING As String = ""
    Private g_sEmailDOAllowed As String = "" 'SY add on 23102020

    Private g_sNCMQUERY_INVQuery As String = ""
    Private g_sNCMQUERY_DLNQuery As String = ""
    Private g_sNCMQUERY_DPIQuery As String = ""
    Private g_sNCMQUERY_RINQuery As String = ""

    Private dsInv As System.Data.DataSet
    Private dsSerInv As System.Data.DataSet

    Private sqlConn As SqlConnection
    Private sqlComm As SqlCommand
    Private da As SqlDataAdapter

#End Region

#Region "Intialize Application"
    Public Sub New()
        Try

        Catch ex As Exception
            MsgBox("[NCM_INV].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        Try
            If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_INV.srf") Then
                oForm = SBO_Application.Forms.Item("NCM_INV")
                oMtrx = oForm.Items.Item("mxList").Specific

                g_bSelectAll = True
                g_sLastFolder = ""
                SetupChooseFromList()
                g_sCOMPANYSETTING = "GENERIC"
                g_sSelectedDocType = "ARINV"

                'SY add on 23102020
                GetUserAuthorization()
                AddDataSource()
                oForm.Visible = True

                ' ================================================================================================= 
                Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim sRec As String = ""
                Dim sQType As String = ""
                sRec = "  SELECT TOP 1 IFNULL(""U_COMPANY"",'')  FROM ""@NCM_SETTING"" "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sRec)
                If oRec.RecordCount > 0 Then
                    g_sCOMPANYSETTING = oRec.Fields.Item(0).Value.ToString.Trim
                End If

                g_sNCMQUERY_INVQuery = ""
                g_sNCMQUERY_DLNQuery = ""
                g_sNCMQUERY_DPIQuery = ""
                g_sNCMQUERY_RINQuery = ""

                sRec = " SELECT ""U_Type"", ""U_Query"" FROM ""@NCM_QUERY"" "
                sRec &= " WHERE ""U_Type"" IN ('NCM_IRP_DLN','NCM_IRP_INV','NCM_IRP_RIN','NCM_IRP_DPI') "
                sRec &= " ORDER BY ""U_Type"" "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sRec)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    While Not oRec.EoF
                        sQType = oRec.Fields.Item("U_Type").Value.ToString.Trim
                        Select Case sQType
                            Case "NCM_IRP_DLN"
                                g_sNCMQUERY_DLNQuery = oRec.Fields.Item("U_Query").Value.ToString.Trim
                            Case "NCM_IRP_INV"
                                g_sNCMQUERY_INVQuery = oRec.Fields.Item("U_Query").Value.ToString.Trim
                            Case "NCM_IRP_RIN"
                                g_sNCMQUERY_RINQuery = oRec.Fields.Item("U_Query").Value.ToString.Trim
                            Case "NCM_IRP_DPI"
                                g_sNCMQUERY_DPIQuery = oRec.Fields.Item("U_Query").Value.ToString.Trim
                        End Select
                        oRec.MoveNext()
                    End While
                End If

                oRec = Nothing
                ' ================================================================================================= 
                g_sOINV_Query = "  SELECT CAST(T1.""DocNum"" AS NVARCHAR(10)) ""DocNum"" , CAST(T1.""DocEntry"" AS NVARCHAR(10)) ""DocEntry"", T1.""DocType"", T1.""CANCELED"", T1.""DocStatus"", "
                g_sOINV_Query &= " T1.""DocDate"", T1.""TaxDate"", T1.""DocDueDate"", T1.""CardCode"", T1.""CardName"", "
                g_sOINV_Query &= " T1.""NumAtCard"", T1.""DocCur"", T1.""DocRate"", T1.""DocTotal"", T1.""DocTotalFC"", T1.""DocTotalSy"", "
                g_sOINV_Query &= " T1.""PaidToDate"", T1.""PaidFC"", T1.""Comments"", T1.""JrnlMemo"", T1.""SlpCode"", "
                g_sOINV_Query &= " T5.""Memo"", T5.""SlpName"", T1.""Footer"", T1.""PayToCode"", "
                g_sOINV_Query &= " T1.""Address"", T1.""Address2"", T1.""CntctCode"", T2.""Name"" ""CntctName"", "
                g_sOINV_Query &= " T1.""Series"" ""SeriesCode"", T6.""SeriesName"", T1.""TotalExpns"", T1.""TotalExpFC"", T1.""TotalExpSC"", "
                g_sOINV_Query &= " IFNULL(T1.""Project"",'') ""Project"", IFNULL(T7.""PrjName"",'') ""PrjName"", "
                g_sOINV_Query &= " T1.""PIndicator"", T1.""PaidSum"", T1.""PaidSumFc"", T1.""OwnerCode"", "
                g_sOINV_Query &= " CASE WHEN IFNULL(T8.""lastName"",'') = '' THEN '' ELSE T8.""lastName"" || N', ' END || "
                g_sOINV_Query &= " CASE WHEN IFNULL(T8.""firstName"",'') = '' THEN '' ELSE T8.""firstName"" || N' ' END || "
                g_sOINV_Query &= " CASE WHEN IFNULL(T8.""middleName"",'') = '' THEN '' ELSE T8.""middleName"" || N'' END ""OwnerName"", "
                g_sOINV_Query &= " T1.""GroupNum"",  IFNULL(T9.""PymntGroup"",'') ""PymntGroup"", T1.""Printed"", "
                g_sOINV_Query &= " T3.""GroupCode"",  "
                g_sOINV_Query &= " IFNULL(T1.""U_INCOTerm"",'') ""U_INCOTerm"", "
                g_sOINV_Query &= " IFNULL(T1.""U_SampleInv"",'') ""U_SampleInv"", "
                g_sOINV_Query &= " IFNULL(T1.""U_Deliverypay"",'') ""U_Deliverypay"", "
                g_sOINV_Query &= " IFNULL(T1.""U_Vessel"",'') ""U_Vessel"", "
                g_sOINV_Query &= " IFNULL(T1.""U_VoyageNo"",'') ""U_VoyageNo"", "
                g_sOINV_Query &= " IFNULL(T1.""U_Location"",'') ""U_Location"", "
                g_sOINV_Query &= " IFNULL(T1.""U_VesselCat"",'') ""U_VesselCat"", "
                g_sOINV_Query &= " IFNULL(T1.""U_Job"",'') ""U_Job"", "
                g_sOINV_Query &= " IFNULL(T1.""U_BPRemarks"",'') ""U_BPRemarks"", "
                g_sOINV_Query &= " IFNULL(T1.""U_ContactPerson"",'') ""U_ContactPerson"", "
                g_sOINV_Query &= " IFNULL(T1.""U_ContractNo"",'') ""U_ContractNo"", "
                g_sOINV_Query &= " IFNULL(T1.""U_ETA"",'19000101') ""U_ETA"", "
                g_sOINV_Query &= " IFNULL(T1.""U_ETD"",'19000101') ""U_ETD"","
                g_sOINV_Query &= " T1.""VatPercent"", T1.""VatSum"", T1.""VatSumFC"", T1.""VatSumSy"", T1.""DiscPrcnt"", T1.""DiscSum"", T1.""DiscSumFC"", T1.""DiscSumSy"", "
                g_sOINV_Query &= " T1.""DpmAmnt"", T1.""DpmAmntFC"", T1.""DpmAmntSC"", T1.""DpmPrcnt"", T1.""DpmAppl"", T1.""DpmApplFc"", T1.""GrosProfit"", T1.""GrosProfFC"" "
                g_sOINV_Query &= " FROM """ & oCompany.CompanyDB & """.""XX_HEADER_XX"" T1 "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCPR"" T2 ON T1.""CntctCode"" = T2.""CntctCode"" AND T1.""CardCode"" = T2.""CardCode"" "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T3 ON T1.""CardCode"" = T3.""CardCode"" "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OSLP"" T5 ON T1.""SlpCode"" = T5.""SlpCode"" "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""NNM1"" T6 ON T1.""Series"" = T6.""Series"" "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OPRJ"" T7 ON T1.""Project"" = T7.""PrjCode"" "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OHEM"" T8 ON T1.""OwnerCode"" = T8.""empID"" "
                g_sOINV_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCTG"" T9 ON T1.""GroupNum"" = T9.""GroupNum"" "

                g_sINV1_Query = "  SELECT T1.""DocEntry"", T1.""LineNum"", '-1' ""LineSeq"", T1.""VisOrder"", T1.""ItemCode"", "
                g_sINV1_Query &= " T1.""Dscription"" ""Description"", T1.""Quantity"", T1.""Price"", T1.""Currency"", "
                g_sINV1_Query &= " T1.""Rate"", T1.""LineTotal"", T1.""TotalFrgn"", T1.""PriceBefDi"", T1.""WhsCode"", "
                g_sINV1_Query &= " T4.""WhsName"", T1.""VatGroup"", T1.""VatPrcnt"", T1.""PriceAfVAT"", T1.""VatSum"", T1.""VatSumFrgn"", "
                g_sINV1_Query &= " IFNULL(T1.""U_ItemDetails"",'') ""U_ItemDetails"", T1.""TreeType"", T1.""BaseDocNum"", T1.""FreeText"", "
                g_sINV1_Query &= " T1.""U_SONo"",  "
                g_sINV1_Query &= " IFNULL(T1.""U_TxnUOM"",'') ""U_TxnUOM"", "
                g_sINV1_Query &= " IFNULL(T1.""U_TxnUOMDesc"",'') ""U_TxnUOMDesc"" "
                g_sINV1_Query &= " FROM """ & oCompany.CompanyDB & """.""XX_ROW_XX"" T1 "
                g_sINV1_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""XX_HEADER_XX"" T2 ON T1.""DocEntry"" = T2.""DocEntry"" "
                g_sINV1_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OITM"" T3 ON T1.""ItemCode"" = T3.""ItemCode"" "
                g_sINV1_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OWHS"" T4 ON T1.""WhsCode""  = T4.""WhsCode"" "

                g_sINV3_Query = "  SELECT T1.""DocEntry"", T1.""ExpnsCode"", T3.""ExpnsName"", "
                g_sINV3_Query &= " T1.""LineTotal"", T1.""TotalFrgn"", T1.""TotalSumSy"",  "
                g_sINV3_Query &= "  T1.""VatGroup"", T1.""VatPrcnt"", T1.""VatSum"", T1.""VatSumFrgn"", T1.""VatSumSy"" "
                g_sINV3_Query &= " FROM """ & oCompany.CompanyDB & """.""XX_ROW3_XX"" T1 "
                g_sINV3_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""XX_HEADER_XX"" T2 ON T1.""DocEntry"" = T2.""DocEntry"" "
                g_sINV3_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OEXD"" T3 ON T1.""ExpnsCode"" = T3.""ExpnsCode"" "

                g_sINV10_Query = "  SELECT T1.""DocEntry"", T1.""AftLineNum"" ""LineNum"", T1.""LineSeq"", IFNULL(T2.""VisOrder"",'-1') ""VisOrder"", "
                g_sINV10_Query &= " T1.""LineText"" ""ItemCode"", "
                g_sINV10_Query &= " T1.""LineText"" ""Description"", 0 ""Quantity"", 0 ""Price"", '' ""Currency"", "
                g_sINV10_Query &= " 1 ""Rate"", 0 ""LineTotal"", 0 ""TotalFrgn"", 0 ""PriceBefDi"", '' ""WhsCode"", "
                g_sINV10_Query &= " '' ""WhsName"", '' ""VatGroup"", 0 ""VatPrcnt"", 0 ""PriceAfVAT"", 0 ""VatSum"", 0 ""VatSumFrgn"", "
                g_sINV10_Query &= " '' ""U_ItemDetails"", "
                g_sINV10_Query &= " '' ""U_TxnUOM"", "
                g_sINV10_Query &= " '' ""U_TxnUOMDesc"" "
                g_sINV10_Query &= " FROM """ & oCompany.CompanyDB & """.""XX_ROWTEXT_XX"" T1 "
                g_sINV10_Query &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""XX_ROW_XX"" T2 ON T1.""DocEntry"" = T2.""DocEntry"" AND T1.""AftLineNum"" = T2.""VisOrder"" "

                SetEnableDisable()
            Else
                Try
                    oForm = SBO_Application.Forms.Item("NCM_INV")
                    If oForm.Visible = False Then
                        oForm.Close()
                    Else
                        oForm.Select()
                    End If
                Catch ex As Exception
                    SBO_Application.StatusBar.SetText("[NCM_INV].[LoadForm] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[NCM_INV].[LoadForm] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddDataSource()
        Try
            Dim oOptn As SAPbouiCOM.OptionBtn
            Dim oLink As SAPbouiCOM.LinkedButton
            Dim oCol As SAPbouiCOM.Column

            With oForm.DataSources.UserDataSources
                .Add("cRow", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 254)
                .Add("cSend", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("cEmail", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
                .Add("cEmailCc", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
                .Add("cDocEntry", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
                .Add("cDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("cGroup", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("cDocDate", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("cDocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("cCardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("cCardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("cCurrency", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
                .Add("cAmount", SAPbouiCOM.BoDataType.dt_SUM)
                .Add("cDocument", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
                .Add("cAtt1", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
                .Add("cAtt2", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
                .Add("cAtt3", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
                .Add("cStatus", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
                .Add("cCDMSStat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10) 'SY add for CDMS Delivered status
                .Add("cDelDate", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20) 'SY add for CDMS - DocDueDate

                .Add("opFrINV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("tbDateFr", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("tbDateTo", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("tbCardFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("tbCardTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("tbNumFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("tbNumTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("tbEntFr", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
                .Add("tbEntTo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
                .Add("tbGroup", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                '.Add("tbDoStatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20) 'SY add on 23102020

            End With

            oOptn = oForm.Items.Item("opFrINV").Specific
            oOptn.DataBind.SetBound(True, "", "opFrINV")
            oOptn = oForm.Items.Item("opFrRIN").Specific
            oOptn.GroupWith("opFrINV")
            oOptn = oForm.Items.Item("opFrDPI").Specific
            oOptn.GroupWith("opFrRIN")
            oOptn = oForm.Items.Item("opFrDO").Specific
            oOptn.GroupWith("opFrDPI")

            oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = 1

            oEdit = oForm.Items.Item("tbGroup").Specific
            oEdit.DataBind.SetBound(True, "", "tbGroup")
            oEdit.ChooseFromListUID = "CFL_BG"
            oEdit.ChooseFromListAlias = "GroupCode"
            oEdit = oForm.Items.Item("tbCardFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbCardFr")
            oEdit.ChooseFromListUID = "CFL_CardFr"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = oForm.Items.Item("tbCardTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbCardTo")
            oEdit.ChooseFromListUID = "CFL_CardTo"
            oEdit.ChooseFromListAlias = "CardCode"

            oEdit = oForm.Items.Item("tbEntFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbEntFr")
            oEdit = oForm.Items.Item("tbEntTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbEntTo")

            oEdit = oForm.Items.Item("tbNumFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbNumFr")
            oEdit.ChooseFromListUID = "CFL_INV_DocNumFr"
            oEdit.ChooseFromListAlias = "DocNum"
            oEdit = oForm.Items.Item("tbNumTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbNumTo")
            oEdit.ChooseFromListUID = "CFL_INV_DocNumTo"
            oEdit.ChooseFromListAlias = "DocNum"

            oEdit = oForm.Items.Item("tbDateFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbDateFr")
            oEdit = oForm.Items.Item("tbDateTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbDateTo")

            oForm.Items.Item("lkCardFr").LinkTo = "tbCardFr"
            oForm.Items.Item("lkCardTo").LinkTo = "tbCardTo"
            oForm.Items.Item("lkEntFr").LinkTo = "tbEntFr"
            oForm.Items.Item("lkEntTo").LinkTo = "tbEntTo"

            oLink = oForm.Items.Item("lkEntFr").Specific
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice
            oLink = oForm.Items.Item("lkEntTo").Specific
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice

            oLink = oForm.Items.Item("lkCardFr").Specific
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner
            oLink = oForm.Items.Item("lkCardTo").Specific
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oCol = oMtrx.Columns.Item("cRow")
            oCol.DataBind.SetBound(True, "", "cRow")
            oCol = oMtrx.Columns.Item("cSend")
            oCol.DataBind.SetBound(True, "", "cSend")
            oCol.ValOn = "1"
            oCol.ValOff = "0"
            oCol = oMtrx.Columns.Item("cEmail")
            oCol.DataBind.SetBound(True, "", "cEmail")
            oCol = oMtrx.Columns.Item("cEmailCC")
            oCol.DataBind.SetBound(True, "", "cEmailCc")
            oCol = oMtrx.Columns.Item("cDocDate")
            oCol.DataBind.SetBound(True, "", "cDocDate")
            oCol = oMtrx.Columns.Item("cDocNum")
            oCol.DataBind.SetBound(True, "", "cDocNum")
            oCol = oMtrx.Columns.Item("cCardName")
            oCol.DataBind.SetBound(True, "", "cCardName")
            oCol = oMtrx.Columns.Item("cCurrency")
            oCol.DataBind.SetBound(True, "", "cCurrency")
            oCol = oMtrx.Columns.Item("cStatus")
            oCol.DataBind.SetBound(True, "", "cStatus")
            oCol = oMtrx.Columns.Item("cGroup")
            oCol.DataBind.SetBound(True, "", "cGroup")
            oCol = oMtrx.Columns.Item("cDocument")
            oCol.DataBind.SetBound(True, "", "cDocument")
            oCol = oMtrx.Columns.Item("cDocType")
            oCol.DataBind.SetBound(True, "", "cDocType")
            oCol = oMtrx.Columns.Item("cAmount")
            oCol.DataBind.SetBound(True, "", "cAmount")

            oCol = oMtrx.Columns.Item("cAtt1")
            oCol.DataBind.SetBound(True, "", "cAtt1")
            oCol = oMtrx.Columns.Item("cAtt2")
            oCol.DataBind.SetBound(True, "", "cAtt2")
            oCol = oMtrx.Columns.Item("cAtt3")
            oCol.DataBind.SetBound(True, "", "cAtt3")

            oCol = oMtrx.Columns.Item("cDocEntry")
            oCol.DataBind.SetBound(True, "", "cDocEntry")
            oLink = oMtrx.Columns.Item("cDocEntry").ExtendedObject

            If g_sEmailDOAllowed = "Y" Then
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes
            Else
                Select Case oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx
                    Case 1
                        oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice
                    Case 2
                        oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                    Case 3
                        oLink.LinkedObject = 203
                    Case 4 'SY add on 23102020
                        oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes
                    Case Else
                        oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice
                End Select
            End If


            oCol = oMtrx.Columns.Item("cCardCode")
            oCol.DataBind.SetBound(True, "", "cCardCode")
            oLink = oMtrx.Columns.Item("cCardCode").ExtendedObject
            oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            If g_sEmailDOAllowed = "Y" Then
                oForm.Items.Item("cbStat").DisplayDesc = True
                oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = 4
                oForm.DataSources.UserDataSources.Item("tbEntFr").ValueEx = 0
                oForm.DataSources.UserDataSources.Item("tbEntTo").ValueEx = 0
                g_sSelectedDocType = "ARDO"
            Else
                oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = 1
                oForm.DataSources.UserDataSources.Item("tbEntFr").ValueEx = 0
                oForm.DataSources.UserDataSources.Item("tbEntTo").ValueEx = 0
                g_sSelectedDocType = "ARINV"
            End If


        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[SetDataSources] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub SetupChooseFromList()
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Try
            oCFLs = oForm.ChooseFromLists

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 2
            oCFLCreation.UniqueID = "CFL_CardFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "2"
            oCFLCreation.UniqueID = "CFL_CardTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 10
            oCFLCreation.UniqueID = "CFL_BG"
            oCFL = oCFLs.Add(oCFLCreation)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            oCon.Alias = "GroupType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Invoice

            oCFLCreation.UniqueID = "CFL_INV_DocNumFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCFLCreation.UniqueID = "CFL_INV_DocNumTo"
            oCFL = oCFLs.Add(oCFLCreation)
            ' ----------------------------------------
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo

            oCFLCreation.UniqueID = "CFL_RIN_DocNumFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCFLCreation.UniqueID = "CFL_RIN_DocNumTo"
            oCFL = oCFLs.Add(oCFLCreation)
            ' ----------------------------------------
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = 203

            oCFLCreation.UniqueID = "CFL_DPI_DocNumFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCFLCreation.UniqueID = "CFL_DPI_DocNumTo"
            oCFL = oCFLs.Add(oCFLCreation)
            ' ----------------------------------------
            'SY add on 23102020
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes

            oCFLCreation.UniqueID = "CFL_DO_DocNumFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oCFLCreation.UniqueID = "CFL_DO_DocNumTo"
            oCFL = oCFLs.Add(oCFLCreation)
            ' ----------------------------------------
        Catch ex As Exception
            Throw New Exception("[NCM_INV].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function ClearFields(ByVal sOption As String) As Boolean
        Try
            Dim oLink As SAPbouiCOM.LinkedButton
            Dim oLinkCol As SAPbouiCOM.LinkedButton

            oMtrx = oForm.Items.Item("mxList").Specific
            oMtrx.Clear()

            With oForm.DataSources.UserDataSources
                .Item("tbEntFr").ValueEx = 0
                .Item("tbEntTo").ValueEx = 0
                .Item("tbNumFr").ValueEx = ""
                .Item("tbNumTo").ValueEx = ""
                .Item("tbDateFr").ValueEx = ""
                .Item("tbDateTo").ValueEx = ""
                .Item("tbCardFr").ValueEx = ""
                .Item("tbCardTo").ValueEx = ""
                .Item("tbGroup").ValueEx = ""
            End With

            oMtrx.Columns.Item("cAtt1").Visible = True
            oMtrx.Columns.Item("cAtt2").Visible = True
            oMtrx.Columns.Item("cAtt3").Visible = True

            Select Case sOption
                Case "ARINV"
                    oEdit = oForm.Items.Item("tbNumFr").Specific
                    oEdit.ChooseFromListUID = "CFL_INV_DocNumFr"
                    oEdit.ChooseFromListAlias = "DocNum"
                    oEdit = oForm.Items.Item("tbNumTo").Specific
                    oEdit.ChooseFromListUID = "CFL_INV_DocNumTo"
                    oEdit.ChooseFromListAlias = "DocNum"

                    oLink = oForm.Items.Item("lkEntFr").Specific
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice
                    oLink = oForm.Items.Item("lkEntTo").Specific
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice
                    oLinkCol = oMtrx.Columns.Item("cDocEntry").ExtendedObject
                    oLinkCol.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True

                Case "ARRIN"
                    oEdit = oForm.Items.Item("tbNumFr").Specific
                    oEdit.ChooseFromListUID = "CFL_RIN_DocNumFr"
                    oEdit.ChooseFromListAlias = "DocNum"
                    oEdit = oForm.Items.Item("tbNumTo").Specific
                    oEdit.ChooseFromListUID = "CFL_RIN_DocNumTo"
                    oEdit.ChooseFromListAlias = "DocNum"

                    oLink = oForm.Items.Item("lkEntFr").Specific
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                    oLink = oForm.Items.Item("lkEntTo").Specific
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo
                    oLinkCol = oMtrx.Columns.Item("cDocEntry").ExtendedObject
                    oLinkCol.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_InvoiceCreditMemo

                Case "ARDPI"
                    oEdit = oForm.Items.Item("tbNumFr").Specific
                    oEdit.ChooseFromListUID = "CFL_DPI_DocNumFr"
                    oEdit.ChooseFromListAlias = "DocNum"
                    oEdit = oForm.Items.Item("tbNumTo").Specific
                    oEdit.ChooseFromListUID = "CFL_DPI_DocNumTo"
                    oEdit.ChooseFromListAlias = "DocNum"

                    oLink = oForm.Items.Item("lkEntFr").Specific
                    oLink.LinkedObject = 203
                    oLink = oForm.Items.Item("lkEntTo").Specific
                    oLink.LinkedObject = 203
                    oLinkCol = oMtrx.Columns.Item("cDocEntry").ExtendedObject
                    oLinkCol.LinkedObject = 203
                Case "ARDO" 'SY add on 23102020
                    oEdit = oForm.Items.Item("tbNumFr").Specific
                    oEdit.ChooseFromListUID = "CFL_DO_DocNumFr"
                    oEdit.ChooseFromListAlias = "DocNum"
                    oEdit = oForm.Items.Item("tbNumTo").Specific
                    oEdit.ChooseFromListUID = "CFL_DO_DocNumTo"
                    oEdit.ChooseFromListAlias = "DocNum"

                    oLink = oForm.Items.Item("lkEntFr").Specific
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes
                    oLink = oForm.Items.Item("lkEntTo").Specific
                    oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes
                    oLinkCol = oMtrx.Columns.Item("cDocEntry").ExtendedObject
                    oLinkCol.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_DeliveryNotes

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True
            End Select

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ClearFields] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    'SY add on 23102020
    Private Sub GetUserAuthorization()
        Try
            '  Select ISNULL(U_EmailDOAllowed, 'N') as U_EmailDOAllowed from OUSR where USER_CODE = 'manager'
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sRec As String = ""
            sRec = "  SELECT TOP 1 IFNULL(""U_EmailDOAllowed"", 'N') AS ""U_EmailDOAllowed"" FROM ""OUSR"" WHERE ""USER_CODE"" = '" & oCompany.UserName & "'"
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sRec)
            If oRec.RecordCount > 0 Then
                g_sEmailDOAllowed = oRec.Fields.Item(0).Value.ToString.Trim
            End If
            oRec = Nothing
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[NCM_INV].[GetUserAuthorization] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    'SY add on 23102020
    Private Sub SetEnableDisable()
        Try
            If g_sEmailDOAllowed = "Y" Then
                oForm.Items.Item("opFrINV").Enabled = False
                oForm.Items.Item("opFrRIN").Enabled = False
                oForm.Items.Item("opFrDPI").Enabled = False
                oForm.Items.Item("opFrDO").Enabled = True
                oForm.Items.Item("cbStat").Enabled = True
                If ClearFields("ARDO") Then
                    oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = "4"
                    oForm.Items.Item("tbGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    g_sSelectedDocType = "ARDO"
                End If
            Else
                oForm.Items.Item("opFrINV").Enabled = True
                oForm.Items.Item("opFrRIN").Enabled = True
                oForm.Items.Item("opFrDPI").Enabled = True
                oForm.Items.Item("opFrDO").Enabled = False
                oForm.Items.Item("cbStat").Enabled = False
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[NCM_INV].[SetEnableDisable] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Logic Function"
    Private Function IsSharedFile(ByVal iLine As Integer, ByVal sGroupcode As String, ByVal sGroupName As String, ByVal sDocType As String, ByVal iGroupCode As Integer, ByVal sLineType As String) As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""
            Dim sParamDocType As String = ""
            Dim sReportFile As String = ""
            Dim sErrorDocType As String = ""
            Dim sErrorDatType As String = ""
            g_sINV_SharedFile = ""

            Select Case sDocType.Trim
                Case "1"
                    sParamDocType = "ARINV"
                    sErrorDocType = "AR Invoice"
                Case "2"
                    sParamDocType = "ARRIN"
                    sErrorDocType = "AR Credit Note"
                Case "3"
                    sParamDocType = "ARDPI"
                    sErrorDocType = "AR Down Payment Invoice"
                Case "4" 'SY add on 23102020
                    sParamDocType = "ARDO"
                    sErrorDocType = "AR Delivery Order"
                Case Else
                    sParamDocType = "ARINV"
                    sErrorDocType = "AR Invoice"
            End Select

            Select Case sLineType
                Case "I"
                    sErrorDatType = "Item"
                Case "S"
                    sErrorDatType = "Service"
            End Select

            sQuery = "  SELECT TOP 1 IFNULL(""U_COMPANY"",'')  FROM ""@NCM_SETTING"" "
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                g_sCOMPANYSETTING = oRec.Fields.Item(0).Value.ToString.Trim
            End If

            sQuery = "  SELECT IFNULL(""U_ReportFile"",'') FROM ""@NCM_LAYOUT"" "
            sQuery &= " WHERE IFNULL(""U_DocType"",'') = '" & sParamDocType & "' AND IFNULL(""U_DatType"",'') = '" & sLineType & "' "

            Select Case sGroupcode
                Case "", "0", "-1"
                    sQuery &= " AND IFNULL(""U_GroupCode"",'') = '' "
                Case Else
                    sQuery &= " AND IFNULL(""U_GroupCode"",'') = '" & iGroupCode & "'"
            End Select

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                sReportFile = oRec.Fields.Item(0).Value
            Else
                sQuery = "  SELECT IFNULL(""U_ReportFile"",'') FROM ""@NCM_LAYOUT"" "
                sQuery &= " WHERE IFNULL(""U_DocType"",'') = '" & sParamDocType & "'"
                sQuery &= " AND IFNULL(""U_DatType"",'') = '" & sLineType & "'"
                sQuery &= " AND IFNULL(""U_GroupCode"",'') = '' "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    sReportFile = oRec.Fields.Item(0).Value
                Else
                End If
            End If

            If sReportFile.Trim <> "" Then
                If (Not File.Exists(sReportFile.Trim)) Then
                    SBO_Application.StatusBar.SetText("[Line " & iLine & "] : Invalid File Path detected in UDT [@NCM_LAYOUT] for " & sErrorDocType & ", type " & sErrorDatType & "(BP Group: " & sGroupcode & " - " & sGroupName & "). Please check. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oRec = Nothing
                    Return False
                End If
                'true for specific group
                g_sINV_SharedFile = sReportFile.Trim
                oRec = Nothing
                Return True
            Else
                'check default customer group code layout
                sQuery = "  SELECT IFNULL(""U_ReportFile"",'') FROM ""@NCM_LAYOUT"" "
                sQuery &= " WHERE IFNULL(""U_DocType"",'') = '" & sParamDocType & "'"
                sQuery &= " AND IFNULL(""U_DatType"",'') = '" & sLineType & "'"
                sQuery &= " AND IFNULL(""U_GroupCode"",'') = '' ORDER BY ""Code"" "
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    sReportFile = oRec.Fields.Item(0).Value

                    If sReportFile.Trim <> "" Then
                        If (Not File.Exists(sReportFile.Trim)) Then
                            SBO_Application.StatusBar.SetText("[Line " & iLine & "] : Invalid Default Layout's File Path detected in UDT [@NCM_LAYOUT] for " & sErrorDocType & ", type " & sErrorDatType & "(BP Group: " & sGroupcode & " - " & sGroupName & "). Please check. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oRec = Nothing
                            Return False
                        End If
                        'true for default layout
                        g_sINV_SharedFile = sReportFile.Trim
                        oRec = Nothing
                        Return True
                    Else
                        SBO_Application.StatusBar.SetText("[Line " & iLine & "] : Blank Default Layout's File Path detected in UDT [@NCM_LAYOUT] for " & sErrorDocType & ", type " & sErrorDatType & "(BP Group: " & sGroupcode & " - " & sGroupName & "). Please check. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Else
                    SBO_Application.StatusBar.SetText("[Line " & iLine & "] : Blank Default Layout's File Path detected in UDT [@NCM_LAYOUT] for " & sErrorDocType & ", type " & sErrorDatType & "(BP Group: " & sGroupcode & " - " & sGroupName & "). Please check. ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            End If

            oRec = Nothing
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[IsSharedFile] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub ListDocument()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim frm As Hydac_FormViewer = New Hydac_FormViewer
            Dim sQuery As String = ""
            Dim sTableHeader As String = ""
            Dim sTableLine As String = ""
            Dim sTableLine3 As String = ""
            Dim sTableLine10 As String = ""

            Dim sCond As String = ""
            Dim sPathFormat As String = ""
            Dim sText As String = ""
            Dim iSelectedOption As Integer = 1
            Dim sTempDirectory As String = ""
            Dim sSelect As String = ""

            Dim ProviderName As String = "System.Data.Odbc"
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            dsInv = New EInvoice
            dtOADM = dsInv.Tables("dtOADM")
            dtEInvoice = dsInv.Tables("dtEInvoice")

            dtOADM.Clear()
            sSelect = "  SELECT TOP 1 T1.""CompnyName"", IFNULL(T1.""CompnyAddr"",'') AS ""CompnyAddr"", "
            sSelect &= " T1.""Phone1"", T1.""Phone1F"", T1.""Phone2"", T1.""Fax"", T1.""FaxF"", T2.""ZipCode"", T2.""ZipCodeF"", "
            sSelect &= " T1.""E_Mail"", T1.""Country"", T1.""RevOffice"", T1.""TaxIdNum"", T1.""FreeZoneNo"", "
            sSelect &= " T1.""PrintHeadr"", T1.""MainCurncy"", T1.""CmpnyAddrF"" "
            sSelect &= " FROM """ & oCompany.CompanyDB & """.""OADM"" T1 "
            sSelect &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""ADM1"" T2 On T1.""Code"" = T2.""Code"" "

            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sSelect
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)

            'dsSerInv = New ESerInvoice
            'ds = New DataSet2
            'dtOINV = ds.Tables("DT_OINV")
            'dtINV1 = ds.Tables("DT_INV1")
            'dtINV3 = ds.Tables("DT_INV3")
            'dtINV10 = ds.Tables("DT_INV10")

            'dtOADM = ds.Tables("DT_OADM")
            'dtADM1 = ds.Tables("DT_ADM1")
            'dtOCRD = ds.Tables("DT_OCRD")
            'dtOCTG = ds.Tables("DT_OCTG")
            'dtOPRJ = ds.Tables("DT_OPRJ")
            'dtOSLP = ds.Tables("DT_OSLP")
            'dtOEXD = ds.Tables("DT_OEXD")
            'dtSRREASON = ds.Tables("DT_NCM_SR_REASON")
            'dtSRADDRESS3 = ds.Tables("DT_SR_ADDRESS3")

            'dtOADM.Clear()
            'dtADM1.Clear()
            'dtOCRD.Clear()
            'dtOCTG.Clear()
            'dtOPRJ.Clear()
            'dtOSLP.Clear()
            'dtOEXD.Clear()
            'dtSRREASON.Clear()
            'dtSRADDRESS3.Clear()

            'sSelect = "  Select TOP 1 ""ZipCode"", ""ZipCodeF"" FROM """ & oCompany.CompanyDB & """.""ADM1"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtADM1)

            'sSelect = " Select ""CardCode"", ""CardName"", ""GroupNum"", ""GroupCode"", ""SlpCode"" FROM """ & oCompany.CompanyDB & """.""OCRD"" WHERE ""CardType"" = 'C' "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtOCRD)

            'sSelect = " SELECT ""GroupNum"", ""PymntGroup"" FROM """ & oCompany.CompanyDB & """.""OCTG"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtOCTG)

            'sSelect = " SELECT ""PrjCode"", ""PrjName"" FROM """ & oCompany.CompanyDB & """.""OPRJ"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtOPRJ)

            'sSelect = " SELECT ""SlpCode"", ""SlpName"", TO_VARCHAR(""Memo"") AS ""Memo"" FROM """ & oCompany.CompanyDB & """.""OSLP"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtOSLP)

            'sSelect = " SELECT ""ExpnsCode"", ""ExpnsName"" FROM """ & oCompany.CompanyDB & """.""OEXD"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtOEXD)

            'sSelect = " SELECT ""Code"", ""Name"" FROM """ & oCompany.CompanyDB & """.""@NCM_SR_REASON"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtSRREASON)

            'sSelect = " SELECT ""Code"", ""Name"", ""U_Address1"", ""U_Address2"", ""U_Address3"", ""U_Tel"", ""U_Fax"" FROM """ & oCompany.CompanyDB & """.""@SR_ADDRESS3"" "
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelect
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtSRADDRESS3)

            Select Case oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx
                Case 1
                    sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARINV\" & oCompany.CompanyDB
                Case 2
                    sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARRIN\" & oCompany.CompanyDB
                Case 3
                    sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARDPI\" & oCompany.CompanyDB
                Case 4 'SY add on 23102020
                    sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARDO\" & oCompany.CompanyDB
                Case Else
                    sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARINV\" & oCompany.CompanyDB
            End Select

            oForm.Items.Item("tbNumTo").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("btList").Enabled = False
            oForm.Items.Item("btSend").Enabled = False

            Select Case oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx
                Case 1
                    iSelectedOption = 1
                    sTableHeader = "OINV"
                    sTableLine = "INV1"
                    sTableLine3 = "INV3"
                    sTableLine10 = "INV10"

                    sPathFormat = "{0}\INV_{1}_{2}.pdf"
                    sText = "AR Invoice"

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True

                Case 2
                    iSelectedOption = 2
                    sTableHeader = "ORIN"
                    sTableLine = "RIN1"
                    sTableLine3 = "RIN3"
                    sTableLine10 = "RIN10"

                    sPathFormat = "{0}\RIN_{1}_{2}.pdf"
                    sText = "AR Credit Note"

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True

                Case 3
                    iSelectedOption = 3
                    sTableHeader = "ODPI"
                    sTableLine = "DPI1"
                    sTableLine3 = "DPI3"
                    sTableLine10 = "DPI10"

                    sPathFormat = "{0}\DPI_{1}_{2}.pdf"
                    sText = "AR Down Payment Invoice"

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True

                    'SY add on 23102020
                Case 4
                    iSelectedOption = 4
                    sTableHeader = "ODLN"
                    sTableLine = "DLN1"
                    sTableLine3 = "DLN3"
                    sTableLine10 = "DLN10"

                    sPathFormat = "{0}\DLN_{1}_{2}.pdf"
                    sText = "AR Delivery Order"

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True

                Case Else
                    iSelectedOption = 1
                    sTableHeader = "OINV"
                    sTableLine = "INV1"
                    sTableLine3 = "INV3"
                    sTableLine10 = "INV10"

                    sPathFormat = "{0}\INV_{1}_{2}.pdf"
                    sText = "AR Invoice"

                    oMtrx.Columns.Item("cAtt1").Visible = True
                    oMtrx.Columns.Item("cAtt2").Visible = True
                    oMtrx.Columns.Item("cAtt3").Visible = True
            End Select

            With oForm.DataSources.UserDataSources
                'document date
                If .Item("tbDateFr").ValueEx.Trim = "" Then
                    If .Item("tbDateTo").ValueEx.Trim <> "" Then
                        sCond &= " AND TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') <= '" & .Item("tbDateTo").ValueEx.Trim & "' "
                    End If
                Else
                    If .Item("tbDateTo").ValueEx.Trim = "" Then
                        sCond &= " AND TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') >= '" & .Item("tbDateFr").ValueEx.Trim & "' "
                    Else
                        sCond &= " AND TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') >= '" & .Item("tbDateFr").ValueEx.Trim & "' AND TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') <= '" & .Item("tbDateTo").ValueEx.Trim & "' "
                    End If
                End If

                'cardcode
                If .Item("tbCardFr").ValueEx.Trim = "" Then
                    If .Item("tbCardTo").ValueEx.Trim <> "" Then
                        sCond &= " AND T1.""CardCode"" <= '" & .Item("tbCardTo").ValueEx.Trim & "' "
                    End If
                Else
                    If .Item("tbCardTo").ValueEx.Trim = "" Then
                        sCond &= " AND T1.""CardCode"" >= '" & .Item("tbCardFr").ValueEx.Trim & "' "
                    Else
                        sCond &= " AND T1.""CardCode"" >= '" & .Item("tbCardFr").ValueEx.Trim & "' AND T1.""CardCode"" <= '" & .Item("tbCardTo").ValueEx.Trim & "' "
                    End If
                End If

                'document number
                If .Item("tbEntFr").ValueEx.Trim = "0" Then
                    If .Item("tbEntTo").ValueEx.Trim <> "0" Then
                        sCond &= " AND T1.""DocEntry"" <= '" & .Item("tbEntTo").ValueEx.Trim & "' "
                    End If
                Else
                    If .Item("tbEntTo").ValueEx.Trim = "0" Then
                        sCond &= " AND T1.""DocEntry"" >= '" & .Item("tbEntFr").ValueEx.Trim & "' "
                    Else
                        sCond &= " AND T1.""DocEntry"" >= '" & .Item("tbEntFr").ValueEx.Trim & "' AND T1.""DocEntry"" <= '" & .Item("tbEntTo").ValueEx.Trim & "' "
                    End If
                End If

                If .Item("tbNumFr").ValueEx.Trim = "" Then
                    If .Item("tbNumTo").ValueEx.Trim <> "" Then
                        sCond &= " AND T1.""DocNum"" <= '" & .Item("tbNumTo").ValueEx.Trim & "' "
                    End If
                Else
                    If .Item("tbNumTo").ValueEx.Trim = "" Then
                        sCond &= " AND T1.""DocNum"" >= '" & .Item("tbNumFr").ValueEx.Trim & "' "
                    Else
                        sCond &= " AND T1.""DocNum"" >= '" & .Item("tbNumFr").ValueEx.Trim & "' AND T1.""DocNum"" <= '" & .Item("tbNumTo").ValueEx.Trim & "' "
                    End If
                End If

                ' customer group
                If .Item("tbGroup").ValueEx.Trim <> "" Then
                    sCond &= " AND T3.""GroupCode"" = '" & .Item("tbGroup").ValueEx.Trim & "' "
                End If

                'SY Add DO status filter + only for DO
                'If .Item("tbDoStatus").ValueEx.Trim <> "" And oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = 4 Then
                If oForm.Items.Item("cbStat").Specific.Value <> "" And .Item("opFrINV").ValueEx = 4 Then
                    sCond &= " AND IFNULL(T1.""U_CDMSDelStat"",'P') = '" & oForm.Items.Item("cbStat").Specific.Selected.Value.Trim & "' "
                End If
            End With

            'SY add on 23102020
            Dim sOptionCheck As Integer = 1
            sOptionCheck = oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx

            Select Case sOptionCheck
                Case 4
                    'Email get from ORDR.U_ContactEmail and ODLN.U_ESignature for Delivery Order
                    If g_sNCMQUERY_DLNQuery = "" Then
                        sQuery = "  SELECT T1.""CardCode"", T1.""CardName"", T1.""DocType"", T1.""DocEntry"", T1.""DocNum"", T1.""DocCur"", T1.""U_ESignature"", "
                        sQuery &= " TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') AS ""DocDate"", "
                        sQuery &= " TO_VARCHAR(T1.""DocDueDate"",'YYYYMMDD') AS ""DocDueDate"", " 'just add
                        sQuery &= " CASE WHEN T1.""DocTotalFC"" = 0 THEN T1.""DocTotal"" ELSE T1.""DocTotalFC"" END AS ""TotalAmt"", "
                        sQuery &= " IFNULL(T1.""U_ContactEmail"", T6.""U_ContactEmail"") ""MailTo"", IFNULL(T3.""GroupCode"",0) ""GroupCode"", IFNULL(T4.""GroupName"",'') ""GroupName"" "
                        sQuery &= " ,IFNULL(T1.""U_CDMSDelStat"",'P') ""CDMSDelStat"" " 'SY add for CDMS DO
                        sQuery &= " FROM """ & oCompany.CompanyDB & """.""" & sTableHeader & """ T1 "
                        sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T3 ON T1.""CardCode""  = T3.""CardCode"" "
                        sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T4 ON T3.""GroupCode"" = T4.""GroupCode"" "
                        sQuery &= " INNER JOIN (Select T1.""BaseEntry"", T0.""DocEntry"" from """ & oCompany.CompanyDB & """.""ODLN"" T0 inner join """ & oCompany.CompanyDB & """.""DLN1"" T1 on T0.""DocEntry"" = T1.""DocEntry"" group by T1.""BaseEntry"", T0.""DocEntry"") T5 on T1.""DocEntry"" = T5.""DocEntry"" "
                        sQuery &= " LEFT OUTER JOIN (Select T0.""U_ContactEmail"", T0.""DocEntry"" from """ & oCompany.CompanyDB & """.""ORDR"" T0 ) T6 on T6.""DocEntry"" = T5.""BaseEntry"" "
                        sQuery &= " WHERE 1 = 1 "
                        sQuery &= sCond
                    Else
                        sQuery = g_sNCMQUERY_DLNQuery
                        sQuery &= sCond
                    End If

                Case Else
                    ' Invoice, DO and DP Invoice
                    Select Case sOptionCheck
                        Case 1
                            ' sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARINV\" & oCompany.CompanyDB
                            If g_sNCMQUERY_INVQuery = "" Then
                                sQuery = "  SELECT T1.""CardCode"", T1.""CardName"", T1.""DocType"", T1.""DocEntry"", T1.""DocNum"", T1.""DocCur"", "
                                sQuery &= " TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') AS ""DocDate"", "
                                sQuery &= " TO_VARCHAR(T1.""DocDueDate"",'YYYYMMDD') AS ""DocDueDate"", " 'just add
                                sQuery &= " CASE WHEN T1.""DocTotalFC"" = 0 THEN T1.""DocTotal"" ELSE T1.""DocTotalFC"" END AS ""TotalAmt"", "
                                sQuery &= " IFNULL(T3.""U_SOA_MailTo"",'') ""MailTo"", IFNULL(T3.""GroupCode"",0) ""GroupCode"", IFNULL(T4.""GroupName"",'') ""GroupName"" "
                                sQuery &= " ,'' ""CDMSDelStat"", '' ""U_ESignature"" " 'Just add
                                sQuery &= " FROM """ & oCompany.CompanyDB & """.""" & sTableHeader & """ T1 "
                                sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T3 ON T1.""CardCode""  = T3.""CardCode"" "
                                sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T4 ON T3.""GroupCode"" = T4.""GroupCode"" "
                                sQuery &= " WHERE 1 = 1 "
                                sQuery &= sCond
                            Else
                                sQuery = g_sNCMQUERY_INVQuery
                                sQuery &= sCond
                            End If

                        Case 2
                            ' sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARRIN\" & oCompany.CompanyDB
                            If g_sNCMQUERY_RINQuery = "" Then
                                sQuery = "  SELECT T1.""CardCode"", T1.""CardName"", T1.""DocType"", T1.""DocEntry"", T1.""DocNum"", T1.""DocCur"", "
                                sQuery &= " TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') AS ""DocDate"", "
                                sQuery &= " TO_VARCHAR(T1.""DocDueDate"",'YYYYMMDD') AS ""DocDueDate"", " 'just add
                                sQuery &= " CASE WHEN T1.""DocTotalFC"" = 0 THEN T1.""DocTotal"" ELSE T1.""DocTotalFC"" END AS ""TotalAmt"", "
                                sQuery &= " IFNULL(T3.""U_SOA_MailTo"",'') ""MailTo"", IFNULL(T3.""GroupCode"",0) ""GroupCode"", IFNULL(T4.""GroupName"",'') ""GroupName"" "
                                sQuery &= " ,'' ""CDMSDelStat"", '' ""U_ESignature"" " 'Just add
                                sQuery &= " FROM """ & oCompany.CompanyDB & """.""" & sTableHeader & """ T1 "
                                sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T3 ON T1.""CardCode""  = T3.""CardCode"" "
                                sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T4 ON T3.""GroupCode"" = T4.""GroupCode"" "
                                sQuery &= " WHERE 1 = 1 "
                                sQuery &= sCond
                            Else
                                sQuery = g_sNCMQUERY_RINQuery
                                sQuery &= sCond
                            End If

                        Case 3
                            ' sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ARDPI\" & oCompany.CompanyDB
                            If g_sNCMQUERY_DPIQuery = "" Then
                                sQuery = "  SELECT T1.""CardCode"", T1.""CardName"", T1.""DocType"", T1.""DocEntry"", T1.""DocNum"", T1.""DocCur"", "
                                sQuery &= " TO_VARCHAR(T1.""DocDate"",'YYYYMMDD') AS ""DocDate"", "
                                sQuery &= " TO_VARCHAR(T1.""DocDueDate"",'YYYYMMDD') AS ""DocDueDate"", " 'just add
                                sQuery &= " CASE WHEN T1.""DocTotalFC"" = 0 THEN T1.""DocTotal"" ELSE T1.""DocTotalFC"" END AS ""TotalAmt"", "
                                sQuery &= " IFNULL(T3.""U_SOA_MailTo"",'') ""MailTo"", IFNULL(T3.""GroupCode"",0) ""GroupCode"", IFNULL(T4.""GroupName"",'') ""GroupName"" "
                                sQuery &= " ,'' ""CDMSDelStat"", '' ""U_ESignature"" " 'Just add
                                sQuery &= " FROM """ & oCompany.CompanyDB & """.""" & sTableHeader & """ T1 "
                                sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T3 ON T1.""CardCode""  = T3.""CardCode"" "
                                sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T4 ON T3.""GroupCode"" = T4.""GroupCode"" "
                                sQuery &= " WHERE 1 = 1 "
                                sQuery &= sCond
                            Else
                                sQuery = g_sNCMQUERY_DPIQuery
                                sQuery &= sCond
                            End If
                    End Select
            End Select

            oRec.DoQuery(sQuery)

            Dim iRow As Integer = 1
            Dim al As New System.Collections.ArrayList()
            Dim sOutputDocNum As String = ""
            Dim sEmailCc As String = GetEmailCCFromUDT()
            Dim ESignPath As String = "" 'SY add 
            Dim iTotalCountDoc As Integer = 0

            If oRec.RecordCount > 0 Then
                iTotalCountDoc = oRec.RecordCount
                oRec.MoveFirst()
                oMtrx.Clear()

                ' set up directory for storing attachment
                Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                If Not di.Exists Then
                    di.Create()
                End If
                ' set up directory for storing attachment

                While Not oRec.EoF
                    SBO_Application.StatusBar.SetText("Populating Document [" & iRow & " out of " & iTotalCountDoc & "] ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    ' Commented out for testing purposes - ES 22/02/2021
                    ' =====================================================
                    'If g_sCOMPANYSETTING = "GENERIC" Then
                    '    dtOINV.Clear()
                    '    dtINV1.Clear()
                    '    dtINV3.Clear()
                    '    dtINV10.Clear()
                    'End If
                    ' =====================================================
                    dtEInvoice.Clear()
                    GenerateDataset(dsInv, oRec.Fields.Item("DocEntry").Value.ToString.Trim, oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx, oRec.Fields.Item("DocType").Value.ToString.Trim)

                    With oForm.DataSources.UserDataSources

                        '* --------------------------------------------------------- */
                        sOutputDocNum = oRec.Fields.Item("DocNum").Value 'DocNum
                        g_bINV_SharedFile = IsSharedFile(iRow, oRec.Fields.Item("GroupCode").Value, oRec.Fields.Item("GroupName").Value, iSelectedOption, oRec.Fields.Item("GroupCode").Value, oRec.Fields.Item("DocType").Value)
                        If oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = 4 Then
                            ESignPath = oRec.Fields.Item("U_ESignature").Value
                        End If
                        '* --------------------------------------------------------- */

                        If g_bINV_SharedFile = True Then

                            frm.Text = sText
                            frm.ReportName = ReportName.AR_Invoice
                            frm.DBUsernameViewer = DBUsername
                            frm.DBPasswordViewer = DBPassword
                            frm.IsExport = True
                            frm.INV_ReportDataset = dsInv
                            frm.INV_IsShared = g_bINV_SharedFile
                            frm.INV_ReportFile = g_sINV_SharedFile
                            frm.INV_DocumentType = oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx
                            frm.INV_ExportDocEntry = oRec.Fields.Item("DocEntry").Value
                            frm.INV_CrystalReportExportType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                            frm.INV_CrystalReportExportPath = String.Format(sPathFormat, di.FullName, sOutputDocNum, Now.ToString("ddMMyyyy"))
                            frm.DO_ESignPath = ESignPath 'SY add

                            ' Commented out for testing purposes - ES 22/02/2021
                            ' =====================================================
                            'Select Case g_sCOMPANYSETTING
                            '    Case "GENERIC"
                            '        frm.OpenInvoiceEmailGenericDataset()
                            '    Case Else
                            '        frm.OpenInvoiceEmail()
                            'End Select
                            ' =====================================================

                            ' frm.OpenInvoiceEmail()
                            frm.OpenInvoiceEmailGenericDataset()
                        End If
                        '* -------------------------------------------------------------- *

                        .Item("cRow").ValueEx = iRow
                        .Item("cEmail").ValueEx = oRec.Fields.Item("MailTo").Value
                        .Item("cEmailCc").ValueEx = sEmailCc
                        .Item("cSend").ValueEx = 0
                        .Item("cGroup").ValueEx = oRec.Fields.Item("GroupCode").Value
                        .Item("cCardCode").ValueEx = oRec.Fields.Item("CardCode").Value
                        .Item("cCardName").ValueEx = oRec.Fields.Item("CardName").Value
                        .Item("cDocEntry").ValueEx = oRec.Fields.Item("DocEntry").Value
                        .Item("cDocNum").ValueEx = oRec.Fields.Item("DocNum").Value
                        .Item("cDocDate").ValueEx = oRec.Fields.Item("DocDate").Value
                        .Item("cDocType").ValueEx = oRec.Fields.Item("DocType").Value
                        .Item("cCurrency").ValueEx = oRec.Fields.Item("DocCur").Value
                        .Item("cAmount").ValueEx = oRec.Fields.Item("TotalAmt").Value
                        .Item("cAtt1").ValueEx = ""
                        .Item("cAtt2").ValueEx = ""
                        .Item("cAtt3").ValueEx = ""

                        If g_bINV_SharedFile = True Then
                            .Item("cDocument").ValueEx = String.Format(sPathFormat, di.FullName, oRec.Fields.Item("DocNum").Value, Now.ToString("ddMMyyyy"))
                            .Item("cStatus").ValueEx = "Ready"
                        Else
                            .Item("cDocument").ValueEx = ""
                            .Item("cStatus").ValueEx = "Not Ready"
                        End If

                        'SY add for CDMS DO - for non DO, just blank
                        .Item("cCDMSStat").ValueEx = oRec.Fields.Item("CDMSDelStat").Value
                        .Item("cDelDate").ValueEx = oRec.Fields.Item("DocDueDate").Value
                    End With
                    oMtrx.AddRow()
                    iRow += 1
                    oRec.MoveNext()
                End While
            End If
            oMtrx.AutoResizeColumns()
            oForm.Items.Item("btList").Enabled = True
            oForm.Items.Item("btSend").Enabled = True
            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)

        Catch ex As Exception
            oForm.Items.Item("btList").Enabled = True
            oForm.Items.Item("btSend").Enabled = True
            SBO_Application.StatusBar.SetText("[ListDocument] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub GenerateDataset(ByRef cDataset As DataSet, ByVal sDocEntry As String, ByVal sDocSelectionType As String, ByVal sDocType As String)
        Try
            Dim sSelHeader As String = ""
            Dim sSelRow As String = ""
            Dim ProviderName As String = "System.Data.Odbc"
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Select Case sDocSelectionType.Trim
                Case "1"
                    Select Case sDocType
                        Case "I"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EINVOICE"" ('" & sDocEntry & "')"
                        Case "S"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EINVOICE"" ('" & sDocEntry & "')"
                    End Select
                Case "2"
                    Select Case sDocType
                        Case "I"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_ECREDITMEMO"" ('" & sDocEntry & "')"
                        Case "S"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_ECREDITMEMO"" ('" & sDocEntry & "')"
                    End Select
                Case "3"
                    Select Case sDocType
                        Case "I"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EDPINVOICE"" ('" & sDocEntry & "')"
                        Case "S"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EDPINVOICE"" ('" & sDocEntry & "')"
                    End Select
                Case "4" 'SY add on 23102020
                    Select Case sDocType
                        Case "I"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EDELIVERYORDER"" ('" & sDocEntry & "')"
                        Case "S"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EDELIVERYORDER"" ('" & sDocEntry & "')"
                    End Select
                Case Else
                    Select Case sDocType
                        Case "I"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EINVOICE"" ('" & sDocEntry & "')"
                        Case "S"
                            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_IRP_EINVOICE"" ('" & sDocEntry & "')"
                    End Select
            End Select

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand = dbConn.CreateCommand()
            dtEInvoice = dsInv.Tables("dtEInvoice")
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtEInvoice)

            'sSelHeader = g_sOINV_Query.Replace("XX_HEADER_XX", sInputTableHeader) & " WHERE T1.""DocEntry"" = " & sDocEntry
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelHeader
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtOINV)

            'sSelHeader = g_sINV3_Query.Replace("XX_HEADER_XX", sInputTableHeader).Replace("XX_ROW3_XX", sInputTableRow3) & " WHERE T1.""DocEntry"" = " & sDocEntry
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelHeader
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtINV3)

            'sSelRow = g_sINV1_Query.Replace("XX_HEADER_XX", sInputTableHeader).Replace("XX_ROW_XX", sInputTableRow) & " WHERE T1.""DocEntry"" = " & sDocEntry
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelRow
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtINV1)

            'sSelRow = g_sINV10_Query.Replace("XX_ROW_XX", sInputTableRow).Replace("XX_ROWTEXT_XX", sInputTableRow & "0") & " WHERE T1.""DocEntry"" = " & sDocEntry
            'HANAcmd = dbConn.CreateCommand()
            'HANAcmd.CommandText = sSelRow
            'HANAcmd.ExecuteNonQuery()
            'HANAda.SelectCommand = HANAcmd
            'HANAda.Fill(dtINV10)

            'sSelRow = " SELECT A.* FROM ( "
            'sSelRow &= g_sINV1_Query.Replace("XX_HEADER_XX", sInputTableHeader).Replace("XX_ROW_XX", sInputTableRow)
            'sSelRow &= " WHERE T1.""DocEntry"" = " & sDocEntry
            'sSelRow &= " UNION ALL "
            'sSelRow &= g_sINV10_Query.Replace("XX_ROW_XX", sInputTableRow).Replace("XX_ROWTEXT_XX", sInputTableRow & "0")
            'sSelRow &= " WHERE T1.""DocEntry"" = " & sDocEntry
            'sSelRow &= " ) A "
            'sSelRow &= " ORDER BY A.""VisOrder"", A.""LineSeq"" "

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GenerateDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub PrintDocument()
        Dim sOutput As String = String.Empty
        Try
            oForm.Items.Item("btSend").Enabled = False
            Dim al As New System.Collections.ArrayList
            Dim sDocType As String = ""
            Dim sEmail As String = ""
            Dim sEmailTo As String = ""
            Dim sEmailCc As String = ""
            Dim sAttachFile As String = ""
            Dim s As New clsEmail

            Select Case oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx
                Case 1
                    s.ReportType = "ARINV"
                    sDocType = "ARINV"
                Case 2
                    s.ReportType = "ARRIN"
                    sDocType = "ARRIN"
                Case 3
                    s.ReportType = "ARDPI"
                    sDocType = "ARDPI"
                Case 4
                    s.ReportType = "ARDO"
                    sDocType = "ARDO"
            End Select

            s.GetSetting("SOA")

            For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                sOutput = ""
                oMtrx.GetLineData(i)
                With oForm.DataSources.UserDataSources
                    sEmail = .Item("cEmail").ValueEx.ToString.Trim
                    If .Item("cSend").ValueEx = "1" Then 'if selected
                        If sEmail.Length > 0 Then 'if there is email address
                            If .Item("cStatus").ValueEx <> "Not Ready" Then 'if ready
                                SBO_Application.StatusBar.SetText("[Line " & i & "] Sending email ...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                If Not al.Contains(.Item("cDocEntry").ValueEx) Then
                                    al.Add(.Item("cDocEntry").ValueEx)

                                    Select Case sDocType
                                        Case "ARINV"
                                            If g_sCOMPANYSETTING = "GENERIC" Then
                                                s.IsGeneric = "GENERIC"
                                                s.Attachment = .Item("cDocument").ValueEx

                                                sAttachFile = ""
                                                sAttachFile = .Item("cAtt1").ValueEx.Trim
                                                If sAttachFile.Length > 0 Then
                                                    If IO.File.Exists(sAttachFile) = True Then
                                                        s.Attachment2 = sAttachFile
                                                    Else
                                                        s.Attachment2 = ""
                                                    End If
                                                Else
                                                    s.Attachment2 = ""
                                                End If

                                                sAttachFile = ""
                                                sAttachFile = .Item("cAtt2").ValueEx.Trim
                                                If sAttachFile.Length > 0 Then
                                                    If IO.File.Exists(sAttachFile) = True Then
                                                        s.Attachment3 = sAttachFile
                                                    Else
                                                        s.Attachment3 = ""
                                                    End If
                                                Else
                                                    s.Attachment3 = ""
                                                End If

                                                sAttachFile = ""
                                                sAttachFile = .Item("cAtt3").ValueEx.Trim
                                                If sAttachFile.Length > 0 Then
                                                    If IO.File.Exists(sAttachFile) = True Then
                                                        s.Attachment4 = sAttachFile
                                                    Else
                                                        s.Attachment4 = ""
                                                    End If
                                                Else
                                                    s.Attachment4 = ""
                                                End If

                                            Else
                                                s.IsGeneric = ""
                                                s.Attachment = .Item("cDocument").ValueEx

                                                sAttachFile = ""
                                                sAttachFile = .Item("cAtt1").ValueEx.Trim
                                                If sAttachFile.Length > 0 Then
                                                    If IO.File.Exists(sAttachFile) = True Then
                                                        s.Attachment2 = sAttachFile
                                                    Else
                                                        s.Attachment2 = ""
                                                    End If
                                                Else
                                                    s.Attachment2 = ""
                                                End If

                                                sAttachFile = ""
                                                sAttachFile = .Item("cAtt2").ValueEx.Trim
                                                If sAttachFile.Length > 0 Then
                                                    If IO.File.Exists(sAttachFile) = True Then
                                                        s.Attachment3 = sAttachFile
                                                    Else
                                                        s.Attachment3 = ""
                                                    End If
                                                Else
                                                    s.Attachment3 = ""
                                                End If

                                                sAttachFile = ""
                                                sAttachFile = .Item("cAtt3").ValueEx.Trim
                                                If sAttachFile.Length > 0 Then
                                                    If IO.File.Exists(sAttachFile) = True Then
                                                        s.Attachment4 = sAttachFile
                                                    Else
                                                        s.Attachment4 = ""
                                                    End If
                                                Else
                                                    s.Attachment4 = ""
                                                End If

                                            End If

                                        Case Else
                                            s.Attachment = .Item("cDocument").ValueEx

                                            sAttachFile = ""
                                            sAttachFile = .Item("cAtt1").ValueEx.Trim
                                            If sAttachFile.Length > 0 Then
                                                If IO.File.Exists(sAttachFile) = True Then
                                                    s.Attachment2 = sAttachFile
                                                Else
                                                    s.Attachment2 = ""
                                                End If
                                            Else
                                                s.Attachment2 = ""
                                            End If

                                            sAttachFile = ""
                                            sAttachFile = .Item("cAtt2").ValueEx.Trim
                                            If sAttachFile.Length > 0 Then
                                                If IO.File.Exists(sAttachFile) = True Then
                                                    s.Attachment3 = sAttachFile
                                                Else
                                                    s.Attachment3 = ""
                                                End If
                                            Else
                                                s.Attachment3 = ""
                                            End If

                                            sAttachFile = ""
                                            sAttachFile = .Item("cAtt3").ValueEx.Trim
                                            If sAttachFile.Length > 0 Then
                                                If IO.File.Exists(sAttachFile) = True Then
                                                    s.Attachment4 = sAttachFile
                                                Else
                                                    s.Attachment4 = ""
                                                End If
                                            Else
                                                s.Attachment4 = ""
                                            End If
                                    End Select

                                    ' ===================================================
                                    sEmailTo = .Item("cEmail").ValueEx.Trim
                                    sEmailCc = .Item("cEmailCC").ValueEx.Trim

                                    If sEmailTo.Trim.Length > 0 Then
                                        If sEmailTo.Substring(sEmailTo.Length - 1, 1) = ";" Then
                                            'remove the last ; of email recipient
                                            sEmailTo = sEmailTo.Substring(0, sEmailTo.Length - 1)
                                        End If
                                        s.EmailTo = sEmailTo
                                    End If


                                    If sEmailCc.Trim.Length > 0 Then
                                        If sEmailCc.Substring(sEmailCc.Length - 1, 1) = ";" Then
                                            'remove the last ; of email cc
                                            sEmailCc = sEmailCc.Substring(0, sEmailCc.Length - 1)
                                        End If
                                        'sEmailTo = sEmailTo & ";" & sEmailCc
                                        s.EmailCc = sEmailCc
                                    End If

                                    ' ============================================================================
                                    'SY add for CDMS
                                    Dim sDocT As String = sDocType
                                    Dim dDelDate As Date = Nothing
                                    If .Item("cCDMSStat").ValueEx = "C" Then
                                        sDocT = "DODEL"
                                    ElseIf .Item("cCDMSStat").ValueEx = "P" Then
                                        sDocT = "DOUDEL"
                                        Try
                                            DateTime.TryParse(.Item("cDelDate").ValueEx, dDelDate)
                                        Catch ex As Exception
                                        End Try
                                    End If
                                    ' ============================================================================

                                    If s.SendEmail_INV(sOutput, dDelDate, sDocType, .Item("cDocNum").ValueEx) Then
                                        .Item("cStatus").ValueEx = "Sent"
                                        UpdatePrinted(sDocType, .Item("cDocEntry").ValueEx)
                                    Else
                                        .Item("cStatus").ValueEx = sOutput
                                    End If
                                    ' ============================================================================

                                Else
                                    .Item("cStatus").ValueEx = "Sent"
                                End If
                            Else
                                .Item("cStatus").ValueEx = "Not Ready"
                            End If
                        Else
                            .Item("cStatus").ValueEx = "Skipped"
                        End If
                    Else
                        .Item("cStatus").ValueEx = "Skipped"
                    End If
                    .Item("cEmail").ValueEx = sEmail
                End With
                oMtrx.SetLineData(i)
            Next
            oMtrx.AutoResizeColumns()
            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            s = Nothing
            oForm.Items.Item("btSend").Enabled = True

        Catch ex As Exception
            oForm.Items.Item("btSend").Enabled = True
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("[SendEmail] : " & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub UpdatePrinted(ByVal sInputDocType As String, ByVal sInputDocEntry As String)
        Try
            Dim ErrCode As Integer = 0
            Dim ErrMsg As String = ""

            Select Case sInputDocType
                Case "ARINV"
                    Dim oINV As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    If oINV.GetByKey(sInputDocEntry) Then
                        oINV.Printed = SAPbobsCOM.PrintStatusEnum.psYes
                        If oINV.Update() <> 0 Then
                            oCompany.GetLastError(ErrCode, ErrMsg)
                            SBO_Application.StatusBar.SetText("[UpdateInvoice - Printed] : " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                Case "ARRIN"
                    Dim oRIN As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    If oRIN.GetByKey(sInputDocEntry) Then
                        oRIN.Printed = SAPbobsCOM.PrintStatusEnum.psYes
                        If oRIN.Update() <> 0 Then
                            oCompany.GetLastError(ErrCode, ErrMsg)
                            SBO_Application.StatusBar.SetText("[UpdateCreditNote - Printed] : " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                Case "ARDPI"
                    Dim oDPI As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                    If oDPI.GetByKey(sInputDocEntry) Then
                        oDPI.Printed = SAPbobsCOM.PrintStatusEnum.psYes
                        If oDPI.Update() <> 0 Then
                            oCompany.GetLastError(ErrCode, ErrMsg)
                            SBO_Application.StatusBar.SetText("[UpdateDownPayment - Printed] : " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
                    'SY add on 23102020
                Case "ARDO"
                    Dim oDO As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    If oDO.GetByKey(sInputDocEntry) Then
                        oDO.Printed = SAPbobsCOM.PrintStatusEnum.psYes
                        If oDO.Update() <> 0 Then
                            oCompany.GetLastError(ErrCode, ErrMsg)
                            SBO_Application.StatusBar.SetText("[UpdateDeliveryOrder - Printed] : " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If
            End Select

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[UpdatePrinted] : " & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub OpenFileDialogHere()
        Dim sFile As String = ""
        Dim frm As New Form2
        frm.Show()
        frm.TopMost = True

        If g_sLastFolder.Trim.Length <= 0 Then
            frm.SaveFileDialog1.InitialDirectory = "C:\"
        End If
        frm.SaveFileDialog1.Filter = "All files (*.*)|*.*"
        frm.SaveFileDialog1.RestoreDirectory = True

        If frm.SaveFileDialog1.ShowDialog = DialogResult.OK Then
            sFile = frm.SaveFileDialog1.FileName
            g_sLastFolder = frm.SaveFileDialog1.FileName
            sFile = sFile.Trim

            frm.Close()
            oForm.Select()
            oMtrx.GetLineData(g_iSelectedLine)
            oForm.DataSources.UserDataSources.Item(g_sColUID).ValueEx = sFile
            oMtrx.SetLineData(g_iSelectedLine)
        Else
            frm.Close()
        End If
        '=========================================
        System.Threading.Thread.CurrentThread.Abort()
    End Sub
    Private Function Validate() As Boolean
        Dim oRecord As SAPbobsCOM.Recordset = Nothing
        Dim sOutput As String = String.Empty
        Dim al As New System.Collections.ArrayList
        Dim sEmailTo As String = ""
        Dim sEmailCc As String = ""
        Dim sEmailType As String = ""
        Dim s() As String
        Dim bContinue As Boolean = True

        Try
            'refresh status
            oForm.Freeze(True)
            oForm.Items.Item("tbGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                oMtrx.GetLineData(i)
                oForm.DataSources.UserDataSources.Item("cStatus").ValueEx = "Ready"
                oMtrx.SetLineData(i)
            Next
            oForm.Freeze(False)

            'validate
            For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                sOutput = ""
                oMtrx.GetLineData(i)
                SBO_Application.StatusBar.SetText("[Line " & i & "] : Validating email address...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                With oForm.DataSources.UserDataSources
                    If .Item("cSend").ValueEx = "1" Then
                        ' ===================================================
                        sEmailTo = .Item("cEmail").ValueEx.Trim
                        sEmailCc = .Item("cEmailCC").ValueEx.Trim

                        If sEmailTo.Length > 0 Then
                            If sEmailTo.Substring(sEmailTo.Length - 1, 1) = ";" Then
                                'remove the last ; of email recipient
                                sEmailTo = sEmailTo.Substring(0, sEmailTo.Length - 1)
                            End If

                            If sEmailCc.Trim.Length > 0 Then
                                If sEmailCc.Substring(sEmailCc.Length - 1, 1) = ";" Then
                                    'remove the last ; of email cc
                                    sEmailCc = sEmailCc.Substring(0, sEmailCc.Length - 1)
                                End If

                                sEmailTo = sEmailTo & ";" & sEmailCc
                            End If

                            If sEmailTo.Contains("<") Or sEmailTo.Contains(">") Then
                                .Item("cStatus").ValueEx = "Invalid"
                                bContinue = False

                            Else
                                s = sEmailTo.Split(";")
                                For x As Integer = 0 To s.Length - 1 Step 1
                                    If s(x).ToString.Trim.Length <= 0 Then
                                        .Item("cStatus").ValueEx = "Invalid"
                                        bContinue = False
                                    End If
                                Next
                            End If
                        Else
                            ' IF NO RECIPIENT EMAIL
                            .Item("cStatus").ValueEx = "Skipped"
                        End If
                        ' ===================================================
                    Else
                        ' IF NOT SELECTED
                        .Item("cStatus").ValueEx = "Skipped"
                    End If
                End With
                oMtrx.SetLineData(i)
            Next

            oForm.Freeze(False)
            If bContinue Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("[Validate] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Friend Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If String.Compare(pVal.ItemUID, "mxList", True) = 0 Then
                            If String.Compare(pVal.ColUID, "cDocument", True) = 0 Then
                                BubbleEvent = False
                                oMtrx.GetLineData(pVal.Row)
                                Dim sPath As String = oForm.DataSources.UserDataSources.Item("cDocument").ValueEx
                                sPath = sPath
                                System.Diagnostics.Process.Start(sPath)
                                Return False
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "btList"
                                If oForm.Items.Item("btList").Enabled Then
                                    myThread = New System.Threading.Thread(AddressOf ListDocument)
                                    myThread.SetApartmentState(Threading.ApartmentState.STA)
                                    myThread.Start()
                                Else
                                    BubbleEvent = False
                                    Return False
                                End If

                            Case "btSend"
                                If oForm.Items.Item("btSend").Enabled Then
                                    If oMtrx.VisualRowCount > 0 Then
                                        If Validate() Then
                                            PrintDocument()
                                        End If
                                    Else
                                        SBO_Application.StatusBar.SetText("No documents found in the list. Please click on button 'List' to generate document(s) into the list.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If

                                Else
                                    BubbleEvent = False
                                    Return False
                                End If

                            Case "2"
                                If Not (oForm.Items.Item("btSend").Enabled And oForm.Items.Item("btList").Enabled) Then
                                    SBO_Application.StatusBar.SetText("Process is still running. You cannot quit this module. Please wait until the process is completed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    BubbleEvent = False
                                    Return BubbleEvent
                                End If

                            Case "opFrINV"
                                If g_sSelectedDocType <> "ARINV" Then
                                    If g_sEmailDOAllowed = "N" Then
                                        Dim iReturn As Integer = 0
                                        iReturn = SBO_Application.MessageBox("This will clear all fields and table. Please confirm.", 2, "&Yes", "&No")
                                        If iReturn = 1 Then
                                            If ClearFields("ARINV") Then
                                                oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = "1"
                                                oForm.Items.Item("tbGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                g_sSelectedDocType = "ARINV"
                                            End If
                                        Else
                                            BubbleEvent = False
                                        End If
                                    End If

                                End If

                            Case "opFrRIN"
                                If g_sSelectedDocType <> "ARRIN" Then
                                    If g_sEmailDOAllowed = "N" Then
                                        Dim iReturn As Integer = 0
                                        iReturn = SBO_Application.MessageBox("This will clear all fields and table. Please confirm.", 2, "&Yes", "&No")
                                        If iReturn = 1 Then
                                            If ClearFields("ARRIN") Then
                                                oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = "2"
                                                oForm.Items.Item("tbGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                g_sSelectedDocType = "ARRIN"
                                            End If
                                        Else
                                            BubbleEvent = False
                                        End If
                                    End If

                                End If
                            Case "opFrDPI"
                                If g_sSelectedDocType <> "ARDPI" Then
                                    If g_sEmailDOAllowed = "N" Then
                                        Dim iReturn As Integer = 0
                                        iReturn = SBO_Application.MessageBox("This will clear all fields and table. Please confirm.", 2, "&Yes", "&No")
                                        If iReturn = 1 Then
                                            If ClearFields("ARDPI") Then
                                                oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = "3"
                                                oForm.Items.Item("tbGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                g_sSelectedDocType = "ARDPI"
                                            End If
                                        Else
                                            BubbleEvent = False
                                        End If
                                    End If

                                End If
                            Case "opFrDO" 'SY add on 23102020
                                If g_sSelectedDocType <> "ARDO" Then
                                    If g_sEmailDOAllowed = "Y" Then
                                        Dim iReturn As Integer = 0
                                        iReturn = SBO_Application.MessageBox("This will clear all fields and table. Please confirm.", 2, "&Yes", "&No")
                                        If iReturn = 1 Then
                                            If ClearFields("ARDO") Then
                                                oForm.DataSources.UserDataSources.Item("opFrINV").ValueEx = "4"
                                                oForm.Items.Item("tbGroup").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                g_sSelectedDocType = "ARDO"
                                            End If
                                        Else
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.ItemUID = "mxList" Then
                            If pVal.Row = 0 Then
                                If pVal.ColUID = "cSend" Then
                                   
                                    Select Case g_bSelectAll
                                        Case True
                                            g_bSelectAll = False
                                            For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                                                oMtrx.GetLineData(i)
                                                oForm.DataSources.UserDataSources.Item(pVal.ColUID).ValueEx = "0"
                                                oMtrx.SetLineData(i)
                                            Next
                                        Case False
                                            g_bSelectAll = True
                                            For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                                                oMtrx.GetLineData(i)
                                                oForm.DataSources.UserDataSources.Item(pVal.ColUID).ValueEx = "1"
                                                oMtrx.SetLineData(i)
                                            Next
                                    End Select

                                End If
                            Else
                                Select Case pVal.ColUID
                                    Case "cAtt1", "cAtt2", "cAtt3"
                                        g_sColUID = pVal.ColUID
                                        g_iSelectedLine = pVal.Row

                                        Dim sFile As String = ""
                                        Dim oEdit As SAPbouiCOM.EditText
                                        oEdit = oMtrx.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific
                                        sFile = oEdit.String.Trim

                                        If sFile.Trim.Length <= 0 Then
                                            myThread = New System.Threading.Thread(AddressOf OpenFileDialogHere)
                                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                                            myThread.Start()
                                            myThread.Join()
                                        Else
                                            System.Diagnostics.Process.Start(sFile)
                                        End If

                                End Select
                            End If
                        End If
                End Select
            Else 'After Action
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                        oCFLEvent = pVal
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvent.SelectedObjects
                        If (Not oDataTable Is Nothing) Then
                            Dim sTemp As String = ""
                            Dim sDocNum As String = ""
                            Select Case oCFLEvent.ChooseFromListUID
                                Case "CFL_CardFr"
                                    sTemp = oDataTable.GetValue("CardCode", 0)
                                    oForm.DataSources.UserDataSources.Item("tbCardFr").ValueEx = sTemp
                                    Exit Select
                                Case "CFL_CardTo"
                                    sTemp = oDataTable.GetValue("CardCode", 0)
                                    oForm.DataSources.UserDataSources.Item("tbCardTo").ValueEx = sTemp
                                    Exit Select
                                Case "CFL_BG"
                                    sTemp = oDataTable.GetValue("GroupCode", 0)
                                    oForm.DataSources.UserDataSources.Item("tbGroup").ValueEx = sTemp
                                    Exit Select
                                Case "CFL_INV_DocNumFr", "CFL_DPI_DocNumFr", "CFL_RIN_DocNumFr", "CFL_DO_DocNumFr" 'SY add on 23102020
                                    sTemp = oDataTable.GetValue("DocEntry", 0)
                                    sDocNum = oDataTable.GetValue("DocNum", 0)
                                    oForm.DataSources.UserDataSources.Item("tbEntFr").ValueEx = sTemp
                                    oForm.DataSources.UserDataSources.Item("tbNumFr").ValueEx = sDocNum
                                    Exit Select
                                Case "CFL_INV_DocNumTo", "CFL_DPI_DocNumTo", "CFL_RIN_DocNumTo", "CFL_DO_DocNumTo" 'SY add on 23102020
                                    sTemp = oDataTable.GetValue("DocEntry", 0)
                                    sDocNum = oDataTable.GetValue("DocNum", 0)
                                    oForm.DataSources.UserDataSources.Item("tbEntTo").ValueEx = sTemp
                                    oForm.DataSources.UserDataSources.Item("tbNumTo").ValueEx = sDocNum
                                    Exit Select
                            End Select
                            Return True
                        End If
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            MsgBox("[ItemEvent]" & vbNewLine & ex.Message)
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class