'' © Copyright © 2007-2019, Inecom Pte Ltd, All rights reserved.
'' =============================================================

Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Common
Imports System.IO

Public Class NCM_CHG_LOG_AUDIT

#Region "Common Variables"
    Friend oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCheck As SAPbouiCOM.CheckBox
    Private oItem As SAPbouiCOM.Item
    Private oRecordset As SAPbobsCOM.Recordset
    Private ds As System.Data.DataSet
    Private dt As System.Data.DataTable
    Private objDataRow As System.Data.DataRow

    Private i As Integer
    Private iNoLabel As Integer
    Private sQuery As String = ""
    Private sFrBinLoc As String = ""
    Private sToBinLoc As String = ""
    Private g_StructureFilename As String = ""
    Private g_sReportFilename As String = ""
    Private g_bIsShared As Boolean = False

#End Region

#Region "Property"
#End Region

#Region "Setting Form"
    Public Sub LoadForm()
        Try
            Try
                ' Check if the form is already loaded.
                oForm = SBO_Application.Forms.Item(FRM_CHG_LOG_AUDIT)
                oForm.Select()
                Return
            Catch ex As Exception ' Silence the exception raised by accessing a form that is not loaded.
            End Try
 
            If LoadFromXML("Inecom_SDK_Reporting_Package." & FILE_CHG_LOG_AUDIT) Then ' Loading .srf file
                oForm = SBO_Application.Forms.Item(FRM_CHG_LOG_AUDIT)
                oForm.SupportedModes = -1

                ds = New dsRpt
                DefineUserDataSource()
                oForm.Freeze(False)
                oForm.Visible = True
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadForm] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub DefineUserDataSource()
        Try
            'oForm.Freeze(True)

            oForm.DataSources.UserDataSources.Add("txtUser", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50)
            oEdit = oForm.Items.Item("txtUser").Specific
            oEdit.DataBind.SetBound(True, , "txtUser")
            oForm.DataSources.UserDataSources.Add("txtUName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            oEdit = oForm.Items.Item("txtUName").Specific
            oEdit.DataBind.SetBound(True, , "txtUName")
            oForm.DataSources.UserDataSources.Add("txtDtFrom", SAPbouiCOM.BoDataType.dt_DATE)
            oEdit = oForm.Items.Item("txtDtFrom").Specific
            oEdit.DataBind.SetBound(True, , "txtDtFrom")
            oForm.DataSources.UserDataSources.Add("txtDtTo", SAPbouiCOM.BoDataType.dt_DATE)
            oEdit = oForm.Items.Item("txtDtTo").Specific
            oEdit.DataBind.SetBound(True, , "txtDtTo")

            oForm.DataSources.UserDataSources.Add("chkBPMD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkBPMD").Specific
            oCheck.DataBind.SetBound(True, , "chkBPMD")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkItemMD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkItemMD").Specific
            oCheck.DataBind.SetBound(True, , "chkItemMD")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkSAR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkSAR").Specific
            oCheck.DataBind.SetBound(True, , "chkSAR")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkSQuo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkSQuo").Specific
            oCheck.DataBind.SetBound(True, , "chkSQuo")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkSO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkSO").Specific
            oCheck.DataBind.SetBound(True, , "chkSO")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkDelv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkDelv").Specific
            oCheck.DataBind.SetBound(True, , "chkDelv")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkRet", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkRet").Specific
            oCheck.DataBind.SetBound(True, , "chkRet")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkARDwn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkARDwn").Specific
            oCheck.DataBind.SetBound(True, , "chkARDwn")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkARInv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkARInv").Specific
            oCheck.DataBind.SetBound(True, , "chkARInv")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkARCM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkARCM").Specific
            oCheck.DataBind.SetBound(True, , "chkARCM")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkInv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkInv").Specific
            oCheck.DataBind.SetBound(True, , "chkInv")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkGRec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkGRec").Specific
            oCheck.DataBind.SetBound(True, , "chkGRec")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkGI", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkGI").Specific
            oCheck.DataBind.SetBound(True, , "chkGI")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkInvTR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkInvTR").Specific
            oCheck.DataBind.SetBound(True, , "chkInvTR")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkInvTrn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkInvTrn").Specific
            oCheck.DataBind.SetBound(True, , "chkInvTrn")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkPAP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkPAP").Specific
            oCheck.DataBind.SetBound(True, , "chkPAP")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkPReq", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkPReq").Specific
            oCheck.DataBind.SetBound(True, , "chkPReq")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkPQuo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkPQuo").Specific
            oCheck.DataBind.SetBound(True, , "chkPQuo")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkPO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkPO").Specific
            oCheck.DataBind.SetBound(True, , "chkPO")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkGRPO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkGRPO").Specific
            oCheck.DataBind.SetBound(True, , "chkGRPO")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkGRet", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkGRet").Specific
            oCheck.DataBind.SetBound(True, , "chkGRet")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkAPDP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkAPDP").Specific
            oCheck.DataBind.SetBound(True, , "chkAPDP")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkAPInv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkAPInv").Specific
            oCheck.DataBind.SetBound(True, , "chkAPInv")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"
            oForm.DataSources.UserDataSources.Add("chkAPCM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkAPCM").Specific
            oCheck.DataBind.SetBound(True, , "chkAPCM")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            oForm.DataSources.UserDataSources.Add("chkFin", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkFin").Specific
            oCheck.DataBind.SetBound(True, , "chkFin")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"
            oForm.DataSources.UserDataSources.Add("chkMJe", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            oCheck = oForm.Items.Item("chkMJe").Specific
            oCheck.DataBind.SetBound(True, , "chkMJe")
            oCheck.ValOn = "Y"
            oCheck.ValOff = "N"

            ' set default value
            Dim StringNow As String = Now.Year.ToString & Right("0" & Now.Month.ToString, 2) & Right("0" & Now.Day.ToString, 2)
            oForm.DataSources.UserDataSources.Item("txtDtFrom").ValueEx = StringNow
            oForm.DataSources.UserDataSources.Item("txtDtTo").ValueEx = StringNow

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[DefineUserDataSource] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "PrintReport"
    Private Function IsSharedFileExist() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = "SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.CHANGE_LOG_AUDIT) & "'"
       
            g_sReportFilename = ""
            g_StructureFilename = ""
            g_sReportFilename = GetSharedFilePath(ReportName.CHANGE_LOG_AUDIT)
       
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
            SBO_Application.StatusBar.SetText("[ChangeLog].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub LoadViewer()
        Try
            Dim sUserId As String = oForm.DataSources.UserDataSources.Item("txtUser").ValueEx
            Dim sDateFrom As String = oForm.DataSources.UserDataSources.Item("txtDtFrom").ValueEx
            Dim sDateTo As String = oForm.DataSources.UserDataSources.Item("txtDtTo").ValueEx
            Dim sBPMD As String = oForm.DataSources.UserDataSources.Item("chkBPMD").ValueEx
            Dim sItemMD As String = oForm.DataSources.UserDataSources.Item("chkItemMD").ValueEx
            Dim sSQuo As String = oForm.DataSources.UserDataSources.Item("chkSQuo").ValueEx
            Dim sSO As String = oForm.DataSources.UserDataSources.Item("chkSO").ValueEx
            Dim sDelv As String = oForm.DataSources.UserDataSources.Item("chkDelv").ValueEx
            Dim sReturn As String = oForm.DataSources.UserDataSources.Item("chkRet").ValueEx
            Dim sARDwn As String = oForm.DataSources.UserDataSources.Item("chkARDwn").ValueEx
            Dim sARInv As String = oForm.DataSources.UserDataSources.Item("chkARInv").ValueEx
            Dim sARCM As String = oForm.DataSources.UserDataSources.Item("chkARCM").ValueEx
            Dim sGRec As String = oForm.DataSources.UserDataSources.Item("chkGRec").ValueEx
            Dim sGI As String = oForm.DataSources.UserDataSources.Item("chkGI").ValueEx
            Dim sInvTR As String = oForm.DataSources.UserDataSources.Item("chkInvTR").ValueEx
            Dim sInvTrn As String = oForm.DataSources.UserDataSources.Item("chkInvTrn").ValueEx
            Dim sPReq As String = oForm.DataSources.UserDataSources.Item("chkPReq").ValueEx
            Dim sPQuo As String = oForm.DataSources.UserDataSources.Item("chkPQuo").ValueEx
            Dim sPO As String = oForm.DataSources.UserDataSources.Item("chkPO").ValueEx
            Dim sGRPO As String = oForm.DataSources.UserDataSources.Item("chkGRPO").ValueEx
            Dim sGRet As String = oForm.DataSources.UserDataSources.Item("chkGRet").ValueEx
            Dim sAPDP As String = oForm.DataSources.UserDataSources.Item("chkAPDP").ValueEx
            Dim sAPInv As String = oForm.DataSources.UserDataSources.Item("chkAPInv").ValueEx
            Dim sAPCM As String = oForm.DataSources.UserDataSources.Item("chkAPCM").ValueEx
            Dim sFin As String = oForm.DataSources.UserDataSources.Item("chkFin").ValueEx
            Dim sMJe As String = oForm.DataSources.UserDataSources.Item("chkMJe").ValueEx
            Dim query As String = String.Empty

            SBO_Application.StatusBar.SetText("Loading Report Data...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

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

            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\CLOG_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
            End If
            oRec = Nothing

            ' ===============================================================================
            ' get the folder of CLOG of the current DB Name
            ' set to local
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\CLOG\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\CLOG_" & sCurrDate & "_" & sCurrTime & ".pdf"
            ' ===============================================================================

            Using ds As DataSet = New DatasetLog
                Using dt As System.Data.DataTable = ds.Tables("RPT_CHG_LOG")
                    'For Hana, need call 2 SP per Object

                    dt.Clear()
                    If sBPMD = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading BP Master", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "2", "OCRD")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "2", "OCRD")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sItemMD = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Item Master", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "4", "OITM")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "4", "OITM")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sSQuo = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Sales Quotation", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "23", "OQUT")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "23", "OQUT")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sSO = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Sales Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "17", "ORDR")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "17", "ORDR")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sDelv = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Delivery", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "15", "ODLN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "15", "ODLN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sReturn = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Return", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "16", "ORDN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "16", "ORDN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sARDwn = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading A/R Down Payment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "203", "ODPI")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "203", "ODPI")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sARInv = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading A/R Invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "13", "OINV")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "13", "OINV")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sARCM = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading A/R Credit Memo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "14", "ORIN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "14", "ORIN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sPReq = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Purchase Request", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "1470000113", "OPRQ")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "1470000113", "OPRQ")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sPQuo = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Purchase Quotation", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "540000006", "OPQT")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "540000006", "OPQT")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sPO = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Purchase Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "22", "OPOR")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using

                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "22", "OPOR")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sGRPO = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Goods Return PO", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "20", "OPDN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "20", "OPDN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sGRet = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Goods Return", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "21", "ORPD")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "21", "ORPD")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sAPDP = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading A/P Down Payment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "204", "ODPO")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "204", "ODPO")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sAPInv = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading A/P Invoice", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "18", "OPCH")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "18", "OPCH")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sAPCM = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading A/P Credit Memo", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "19", "ORPC")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "19", "ORPC")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sGRec = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Goods Receipt", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "59", "OIGN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "59", "OIGN")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sGI = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Goods Issue", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "60", "OIGE")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "60", "OIGE")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sInvTR = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Inventory Transfer Request", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "1250000001", "OWTQ")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "1250000001", "OWTQ")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sInvTrn = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Inventory Transfer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "67", "OWTR")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "67", "OWTR")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If

                    If sMJe = "Y" Then
                        SBO_Application.StatusBar.SetText("Loading Manual Journal Entry", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE", sUserId, sDateFrom, sDateTo, "30", "OJDT")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                        query = BuildQueryStr("NCM_SP_CHG_LOG_AUDIT_COMBINE_OBJ", sUserId, sDateFrom, sDateTo, "30", "OJDT")
                        Using da As DbDataAdapter = ExecuteHANACommandToDataAdapter(query)
                            da.Fill(ds, "RPT_CHG_LOG")
                        End Using
                    End If


                    SBO_Application.StatusBar.SetText("Loading Report...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Dim sCmpNm As String = SelectCompanyName()
                    Dim viewer As New Hydac_FormViewer

                    viewer.Text = "Change Log Audit Report"
                    viewer.Name = "Change Log Audit Report"
                    viewer.IsShared = g_bIsShared
                    viewer.SharedReportName = g_sReportFilename
                    viewer.ExportPath = sFinalFileName
                    viewer.ReportName = ReportName.CHANGE_LOG_AUDIT
                    viewer.Dataset = ds
                    viewer.CL_CompanyNm = sCmpNm
                    viewer.CL_GenBy = oCompany.UserName

                    viewer.crViewer.Zoom(100)
                    viewer.TopMost = True

                    If String.IsNullOrEmpty(sUserId) Then
                        viewer.CL_UserID = "ALL"
                    Else
                        viewer.CL_UserID = sUserId
                    End If

                    If Not String.IsNullOrEmpty(sDateFrom) Then
                        viewer.CL_DateFrom = Date.ParseExact(sDateFrom, "yyyyMMdd", Nothing)
                    End If

                    If Not String.IsNullOrEmpty(sDateTo) Then
                        viewer.CL_DateTo = Date.ParseExact(sDateTo, "yyyyMMdd", Nothing)
                    End If

                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            viewer.ClientType = "D"
                        Case SAPbouiCOM.BoClientType.ct_Browser
                            viewer.ClientType = "S"
                    End Select

                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            viewer.ShowDialog()

                        Case SAPbouiCOM.BoClientType.ct_Browser
                            viewer.OPEN_HANADS_CHANGE_LOG()

                            If File.Exists(sFinalFileName) Then
                                SBO_Application.SendFileToBrowser(sFinalFileName)
                            End If
                    End Select


                End Using
            End Using


        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadPrintData] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function BuildQueryStr(ByVal SPName As String, ByVal sUserCode As String, ByVal sDateFrom As String, ByVal sDateTo As String, ByVal sObjRef As String, ByVal sTableName As String) As String
        Dim query As String = String.Empty
        Dim str As String = "CALL """ & oCompany.CompanyDB & """.""" & SPName & """ ( '{0}', '{1}', '{2}', {3}, '{4}' )"
        If (String.IsNullOrEmpty(sUserCode)) Then
            query = String.Format(str, "ALL", sDateFrom, sDateTo, sObjRef, sTableName)

        Else
            query = String.Format(str, sUserCode, sDateFrom, sDateTo, sObjRef, sTableName)
        End If

        Return query
    End Function

    Private Function SelectCompanyName() As String
        Dim sCompnyNm As String = ""
        Dim oRS As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try

            sQuery = "Select ""CompnyName"" From OADM"
            oRS.DoQuery(sQuery)
            oRS.MoveFirst()
            sCompnyNm = oRS.Fields.Item("CompnyName").Value
        Catch ex As Exception
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS)
        End Try
        Return sCompnyNm
    End Function
#End Region

#Region "EventHandlers"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.Before_Action = True Then
            Else
                Select Case pval.EventType
                    'Case SAPbouiCOM.BoEventTypes.et_VALIDATE


                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "btnPrint" Then
                            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                            myThread.SetApartmentState(System.Threading.ApartmentState.STA)
                            myThread.Start()
                        End If
                        If pval.ItemUID = "chkSAR" Then
                            Dim checked As String = oForm.DataSources.UserDataSources.Item("chkSAR").ValueEx
                            If checked = "Y" Then
                                oForm.DataSources.UserDataSources.Item("chkSQuo").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkSO").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkDelv").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkRet").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkARDwn").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkARInv").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkARCM").Value = "Y"
                            Else
                                oForm.DataSources.UserDataSources.Item("chkSQuo").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkSO").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkDelv").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkRet").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkARDwn").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkARInv").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkARCM").Value = "N"

                            End If
                        End If
                        If pval.ItemUID = "chkPAP" Then
                            Dim checked As String = oForm.DataSources.UserDataSources.Item("chkPAP").ValueEx
                            If checked = "Y" Then
                                oForm.DataSources.UserDataSources.Item("chkPReq").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkPQuo").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkPO").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkGRPO").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkGRet").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkAPDP").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkAPInv").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkAPCM").Value = "Y"
                            Else
                                oForm.DataSources.UserDataSources.Item("chkPReq").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkPQuo").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkPO").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkGRPO").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkGRet").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkAPDP").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkAPInv").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkAPCM").Value = "N"
                            End If
                        End If
                        If pval.ItemUID = "chkInv" Then
                            Dim checked As String = oForm.DataSources.UserDataSources.Item("chkInv").ValueEx
                            If checked = "Y" Then
                                oForm.DataSources.UserDataSources.Item("chkGRec").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkGI").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkInvTR").Value = "Y"
                                oForm.DataSources.UserDataSources.Item("chkInvTrn").Value = "Y"
                            Else
                                oForm.DataSources.UserDataSources.Item("chkGRec").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkGI").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkInvTR").Value = "N"
                                oForm.DataSources.UserDataSources.Item("chkInvTrn").Value = "N"

                            End If
                        End If
                        If pval.ItemUID = "chkFin" Then
                            Dim checked As String = oForm.DataSources.UserDataSources.Item("chkFin").ValueEx
                            If checked = "Y" Then
                                oForm.DataSources.UserDataSources.Item("chkMJe").Value = "Y"
                            Else
                                oForm.DataSources.UserDataSources.Item("chkMJe").Value = "N"
                            End If

                        End If
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        Dim sCFL_ID As String
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oCFLEvento = pval
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        oDataTable = oCFLEvento.SelectedObjects

                        Select Case pval.ItemUID

                            Case "txtUser"

                                Try

                                    Dim sCusCode As String = oDataTable.GetValue("USER_CODE", 0).ToString
                                    Dim sCusName As String = oDataTable.GetValue("U_NAME", 0).ToString


                                    oForm.DataSources.UserDataSources.Item("txtUser").Value = sCusCode
                                    oForm.DataSources.UserDataSources.Item("txtUName").Value = sCusName


                                Catch ex As Exception

                                End Try



                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        If pval.ItemUID = "txtUser" Then
                            Dim usercode As String = oForm.DataSources.UserDataSources.Item("txtUser").Value
                            If String.IsNullOrEmpty(usercode) Then
                                'oEdit = oForm.DataSources.UserDataSources.Item("txtUName")
                                oForm.DataSources.UserDataSources.Item("txtUName").Value = String.Empty
                                'oForm.Items.Item("txtUName").Enabled = False
                                'Else
                                'oForm.Items.Item("txtUName").Enabled = True
                            End If
                        End If

                End Select
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
    Public Function MenuEvent(ByRef pval As SAPbouiCOM.MenuEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = True Then
            Else
            End If
            Return BubbleEvent

        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
