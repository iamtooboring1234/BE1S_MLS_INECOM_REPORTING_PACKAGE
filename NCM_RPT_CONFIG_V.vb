'' © Copyright © 2007-2020, Inecom Pte Ltd, All rights reserved.
'' =============================================================
Option Strict Off
Option Explicit On

Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class NCM_RPT_CONFIG_V

#Region "Global Variables"
    Private oFormCFG As Form
    Private oMtrxBSC As Matrix
    Private oMtrxADV As Matrix

    Private sErrMsg As String
    Private lErrCode As Integer
    Private g_sUpd, g_sVie, g_sPrt As String
    Private g_sMatrix As String
    Private g_iPaneLvl As Integer = 1
#End Region

#Region "Constructors"
    Public Sub New()
        MyBase.new()
    End Sub
#End Region

#Region "General Functions"
    Friend Sub LoadForm()
        Dim oFldr As SAPbouiCOM.Folder
        Dim oCols As SAPbouiCOM.Columns
        Dim oColn As SAPbouiCOM.Column
        Dim oCbox As SAPbouiCOM.ComboBox
        Dim oEdit As SAPbouiCOM.EditText

        Try
            If LoadFromXML("Inecom_SDK_Reporting_Package." & FRM_RPT_CONFIG & ".srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading form...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning)
                oFormCFG = SBO_Application.Forms.Item(FRM_RPT_CONFIG)
                oFormCFG.Title = "Reporting Package Configuration"
                oFormCFG.EnableMenu(MenuID.Find, False)
                oFormCFG.EnableMenu(MenuID.Add, False)
                oFormCFG.EnableMenu(MenuID.Add_Row, False)
                oFormCFG.EnableMenu(MenuID.Delete_Row, False)
                oFormCFG.Items.Item("flOne").AffectsFormMode = False
                oFormCFG.Items.Item("flTwo").AffectsFormMode = False

                With oFormCFG.DataSources.UserDataSources
                    .Add("uFLONE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("uFLTWO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("uFLTHREE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("uFLFOUR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("tbPercent", BoDataType.dt_PERCENT)

                    .Add("xRow", BoDataType.dt_SHORT_NUMBER, 10)
                    .Add("xRptCode", BoDataType.dt_SHORT_TEXT, 254)
                    .Add("xRptName", BoDataType.dt_SHORT_TEXT, 254)
                    .Add("xIncluded", BoDataType.dt_SHORT_TEXT, 254)
                    .Add("xFilePath", BoDataType.dt_SHORT_TEXT, 254)

                    .Add("yRow", BoDataType.dt_SHORT_NUMBER, 10)
                    .Add("yRptCode", BoDataType.dt_SHORT_TEXT, 254)
                    .Add("yRptName", BoDataType.dt_SHORT_TEXT, 254)
                    .Add("yIncluded", BoDataType.dt_SHORT_TEXT, 254)
                    .Add("yFilePath", BoDataType.dt_SHORT_TEXT, 254)

                    .Add("cbSOAPrt", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cbGSTCurr", BoDataType.dt_SHORT_TEXT, 30)
                    .Add("cbPVDat", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cbPVInv", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cbIRDat", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cbIRInv", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cbEmailTyp", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("tbNotes", BoDataType.dt_LONG_TEXT)
                    .Add("cbInvType", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("tbInvText", BoDataType.dt_LONG_TEXT)
                    .Add("cbRinType", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("tbRinText", BoDataType.dt_LONG_TEXT)
                    .Add("cbDpiType", BoDataType.dt_SHORT_TEXT, 1)
                    .Add("tbDpiText", BoDataType.dt_LONG_TEXT)
                    .Add("cbCardOp", BoDataType.dt_SHORT_TEXT, 1)

                End With

                oFldr = oFormCFG.Items.Item("flOne").Specific
                oFldr.DataBind.SetBound(True, "", "uFLONE")
                oFldr = oFormCFG.Items.Item("flTwo").Specific
                oFldr.DataBind.SetBound(True, "", "uFLTWO")
                oFldr.GroupWith("flOne")
                oFldr = oFormCFG.Items.Item("flThree").Specific
                oFldr.DataBind.SetBound(True, "", "uFLTHREE")
                oFldr.GroupWith("flTwo")
                oFldr = oFormCFG.Items.Item("flFour").Specific
                oFldr.DataBind.SetBound(True, "", "uFLFOUR")
                oFldr.GroupWith("flThree")

                Dim oRec As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                Dim sRec As String = ""
                Dim sLoc As String = ""

                sRec = "SELECT T0.""MainCurncy"" FROM OADM T0"
                oRec.DoQuery(sRec)
                If oRec.RecordCount > 0 Then
                    sLoc = oRec.Fields.Item(0).Value
                End If

                sRec = "SELECT T0.""CurrCode"", T0.""CurrName"" FROM OCRN T0 ORDER BY T0.""CurrCode"""
                oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRec.DoQuery(sRec)

                oCbox = oFormCFG.Items.Item("cbGSTCurr").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbGSTCurr")
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    While Not oRec.EoF
                        oCbox.ValidValues.Add(oRec.Fields.Item(0).Value, oRec.Fields.Item(1).Value)
                        oRec.MoveNext()
                    End While
                    oCbox.Select(sLoc, BoSearchKey.psk_ByValue)
                End If
                oRec = Nothing

                oCbox = oFormCFG.Items.Item("cbPVInv").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbPVInv")
                oCbox.ValidValues.Add("Y", "Yes")
                oCbox.ValidValues.Add("N", "No")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbIRInv").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbIRInv")
                oCbox.ValidValues.Add("Y", "Yes")
                oCbox.ValidValues.Add("N", "No")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbPVDat").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbPVDat")
                oCbox.ValidValues.Add("Y", "Document Date")
                oCbox.ValidValues.Add("N", "Posting Date")
                oCbox.ValidValues.Add("D", "Due Date")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbIRDat").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbIRDat")
                oCbox.ValidValues.Add("Y", "Document Date")
                oCbox.ValidValues.Add("N", "Posting Date")
                oCbox.ValidValues.Add("D", "Due Date")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbSOAPrt").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbSOAPrt")
                oCbox.ValidValues.Add("2", "Preview & Email")
                oCbox.ValidValues.Add("3", "Preview All, Preview Non-Email & Email")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                oFormCFG.Items.Item("cbSOAPrt").DisplayDesc = True

                ' =============================================================
                oCbox = oFormCFG.Items.Item("cbCardOp").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbCardOp")
                oCbox.ValidValues.Add("C", "BP Code")
                oCbox.ValidValues.Add("N", "BP Name")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oEdit = oFormCFG.Items.Item("tbNotes").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "tbNotes")

                oCbox = oFormCFG.Items.Item("cbEmailTyp").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbEmailTyp")
                oCbox.ValidValues.Add("H", "HTML")
                oCbox.ValidValues.Add("P", "Plain Text")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbInvType").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbInvType")
                oCbox.ValidValues.Add("H", "HTML")
                oCbox.ValidValues.Add("P", "Plain Text")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbRinType").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbRinType")
                oCbox.ValidValues.Add("H", "HTML")
                oCbox.ValidValues.Add("P", "Plain Text")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormCFG.Items.Item("cbDpiType").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cbDpiType")
                oCbox.ValidValues.Add("H", "HTML")
                oCbox.ValidValues.Add("P", "Plain Text")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oEdit = oFormCFG.Items.Item("tbInvText").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "tbInvText")
                oEdit = oFormCFG.Items.Item("tbRinText").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "tbRinText")
                oEdit = oFormCFG.Items.Item("tbDpiText").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "tbDpiText")

                ' =============================================================
                oMtrxBSC = oFormCFG.Items.Item("mxBSC").Specific
                oCols = oMtrxBSC.Columns
                oColn = oCols.Item("cRow")
                oColn.DataBind.SetBound(True, "", "xRow")
                oColn = oCols.Item("cRptCode")
                oColn.DataBind.SetBound(True, "", "xRptCode")
                oColn = oCols.Item("cRptName")
                oColn.DataBind.SetBound(True, "", "xRptName")
                oColn = oCols.Item("cIncluded")
                oColn.DataBind.SetBound(True, "", "xIncluded")
                oColn.ValOff = "N"
                oColn.ValOn = "Y"
                oColn = oCols.Item("cFilePath")
                oColn.DataBind.SetBound(True, "", "xFilePath")

                oMtrxADV = oFormCFG.Items.Item("mxADV").Specific
                oCols = oMtrxADV.Columns
                oColn = oCols.Item("cRow")
                oColn.DataBind.SetBound(True, "", "yRow")
                oColn = oCols.Item("cRptCode")
                oColn.DataBind.SetBound(True, "", "yRptCode")
                oColn = oCols.Item("cRptName")
                oColn.DataBind.SetBound(True, "", "yRptName")
                oColn = oCols.Item("cIncluded")
                oColn.DataBind.SetBound(True, "", "yIncluded")
                oColn.ValOff = "N"
                oColn.ValOn = "Y"
                oColn = oCols.Item("cFilePath")
                oColn.DataBind.SetBound(True, "", "yFilePath")

                g_iPaneLvl = 1
                PopulateReport()
                ''-------------------------------------------------------------------------------------------
                oMtrxBSC.AutoResizeColumns()
                oMtrxADV.AutoResizeColumns()

                oFormCFG.Mode = BoFormMode.fm_OK_MODE
                oFormCFG.Items.Item("flOne").Click(BoCellClickType.ct_Regular)
                oFormCFG.PaneLevel = 1
                oFormCFG.Visible = True

                SBO_Application.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None)
            Else
                ' Loading .srf file failed most likely it is because the form is already opened
                Try
                    oFormCFG = SBO_Application.Forms.Item(FRM_RPT_CONFIG)
                    If oFormCFG.Visible Then
                        oFormCFG.Select()
                    Else
                        oFormCFG.Close()
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadForm] : " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub PopulateReport()
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            Dim iCount As Integer = 1
            Dim sDocType As String = ""

            sQuery = "  SELECT  IFNULL(T0.""U_GSTCURR"",''), IFNULL(T0.""U_INVDETAIL"",'N'), "
            sQuery &= "         IFNULL(T0.""U_TAXDATE"",'Y'), IFNULL(T0.""U_IRAINVDETAIL"",'N'), "
            sQuery &= "         IFNULL(T0.""U_IRATAXDATE"",'Y'), "
            sQuery &= "         IFNULL(T0.""U_EmailType"",'H'), "
            sQuery &= "         IFNULL(T0.""U_PlainText"",''),     IFNULL(T0.""U_InvEmailType"",'H'), IFNULL(T0.""U_InvPlainText"",''), "
            sQuery &= "         IFNULL(T0.""U_RinEmailType"",'H'), IFNULL(T0.""U_RinPlainText"",''),  IFNULL(T0.""U_DpiEmailType"",'H'), "
            sQuery &= "         IFNULL(T0.""U_DpiPlainText"",''),  IFNULL(T0.""U_CardOption"",'C') "
            sQuery &= " FROM    ""@NCM_NEW_SETTING"" T0"

            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            With oFormCFG.DataSources.UserDataSources
                If oRecord.RecordCount > 0 Then
                    oRecord.MoveFirst()
                    If oRecord.Fields.Item(0).Value <> "" Then
                        .Item("cbGSTCurr").ValueEx = oRecord.Fields.Item(0).Value
                    End If
                    .Item("cbPVInv").ValueEx = oRecord.Fields.Item(1).Value
                    .Item("cbPVDat").ValueEx = oRecord.Fields.Item(2).Value
                    .Item("cbIRInv").ValueEx = oRecord.Fields.Item(3).Value
                    .Item("cbIRDat").ValueEx = oRecord.Fields.Item(4).Value
                    .Item("cbEmailTyp").ValueEx = oRecord.Fields.Item(5).Value
                    .Item("tbNotes").ValueEx = oRecord.Fields.Item(6).Value
                    .Item("cbInvType").ValueEx = oRecord.Fields.Item(7).Value
                    .Item("tbInvText").ValueEx = oRecord.Fields.Item(8).Value
                    .Item("cbRinType").ValueEx = oRecord.Fields.Item(9).Value
                    .Item("tbRinText").ValueEx = oRecord.Fields.Item(10).Value
                    .Item("cbDpiType").ValueEx = oRecord.Fields.Item(11).Value
                    .Item("tbDpiText").ValueEx = oRecord.Fields.Item(12).Value
                    .Item("cbCardOp").ValueEx = oRecord.Fields.Item(13).Value
                Else
                    .Item("cbGSTCurr").ValueEx = oCompany.GetCompanyService.GetAdminInfo.LocalCurrency
                    .Item("cbPVInv").ValueEx = "N"
                    .Item("cbPVDat").ValueEx = "Y"
                    .Item("cbIRInv").ValueEx = "N"
                    .Item("cbIRDat").ValueEx = "Y"
                    .Item("cbEmailTyp").ValueEx = "H"
                    .Item("tbNotes").ValueEx = ""
                    .Item("cbInvType").ValueEx = "H"
                    .Item("tbInvText").ValueEx = ""
                    .Item("cbRinType").ValueEx = "H"
                    .Item("tbRinText").ValueEx = ""
                    .Item("cbDpiType").ValueEx = "H"
                    .Item("tbDpiText").ValueEx = ""
                    .Item("cbCardOp").ValueEx = "C"
                End If
            End With

            ' NEW PARAMETER - U_SOAPRTOPT
            ' ==================================================================================
            Try
                sQuery = "  SELECT TOP 1 IFNULL(""U_SOAPRTOPT"",'2') FROM ""@NCM_NEW_SETTING"" "
                oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)
                If oRecord.RecordCount > 0 Then
                    oFormCFG.DataSources.UserDataSources.Item("cbSOAPrt").ValueEx = oRecord.Fields.Item(0).Value.ToString.Trim
                Else
                    oFormCFG.DataSources.UserDataSources.Item("cbSOAPrt").ValueEx = "2"
                End If
            Catch ex As Exception
                oFormCFG.DataSources.UserDataSources.Item("cbSOAPrt").ValueEx = "2"
            End Try
            ' ==================================================================================

            sQuery = "  SELECT   T0.""RPTCODE"", T0.""RPTNAME"", T0.""INCLUDED"", T0.""FILEPATH"" "
            sQuery &= " FROM     ""@NCM_RPT_CONFIG"" T0 "
            sQuery &= " WHERE    T0.""RPTTYPE"" = 'BSC' AND IFNULL(T0.""INCLUDED"",'N') = 'Y' "
            sQuery &= " ORDER BY T0.""RPTNAME"" "
            oMtrxBSC.Clear()
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                While Not oRecord.EoF
                    With oFormCFG.DataSources.UserDataSources
                        .Item("xRow").ValueEx = iCount
                        .Item("xRptCode").ValueEx = oRecord.Fields.Item("RPTCODE").Value
                        .Item("xRptName").ValueEx = oRecord.Fields.Item("RPTNAME").Value
                        .Item("xIncluded").ValueEx = oRecord.Fields.Item("INCLUDED").Value
                        .Item("xFilePath").ValueEx = oRecord.Fields.Item("FILEPATH").Value
                    End With
                    iCount += 1
                    oMtrxBSC.AddRow()
                    oRecord.MoveNext()
                End While
            End If

            sQuery = "  SELECT   T0.""RPTCODE"", T0.""RPTNAME"", T0.""INCLUDED"", T0.""FILEPATH"" "
            sQuery &= " FROM     ""@NCM_RPT_CONFIG"" T0 "
            sQuery &= " WHERE    T0.""RPTTYPE"" = 'ADV' AND IFNULL(T0.""INCLUDED"",'N') = 'Y'"
            sQuery &= " ORDER BY T0.""RPTNAME"" "
            oMtrxADV.Clear()
            iCount = 1
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                While Not oRecord.EoF
                    With oFormCFG.DataSources.UserDataSources
                        .Item("yRow").ValueEx = iCount
                        .Item("yRptCode").ValueEx = oRecord.Fields.Item("RPTCODE").Value
                        .Item("yRptName").ValueEx = oRecord.Fields.Item("RPTNAME").Value
                        .Item("yIncluded").ValueEx = oRecord.Fields.Item("INCLUDED").Value
                        .Item("yFilePath").ValueEx = oRecord.Fields.Item("FILEPATH").Value
                    End With
                    iCount += 1
                    oMtrxADV.AddRow()
                    oRecord.MoveNext()
                End While
            End If
            oRecord = Nothing
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PopulateReport] : " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function UpdateReport() As Boolean
        Try
            Dim sUpdate As String = ""
            Dim oUpdate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            With oFormCFG.DataSources.UserDataSources
                sUpdate = "  UPDATE ""@NCM_NEW_SETTING"" "
                sUpdate &= " SET  ""U_GSTCURR"" = '" & .Item("cbGSTCurr").ValueEx & "', "
                sUpdate &= " ""U_INVDETAIL""    = '" & .Item("cbPVInv").ValueEx & "', "
                sUpdate &= " ""U_TAXDATE""      = '" & .Item("cbPVDat").ValueEx & "', "
                sUpdate &= " ""U_IRAINVDETAIL"" = '" & .Item("cbIRInv").ValueEx & "', "
                sUpdate &= " ""U_IRATAXDATE""   = '" & .Item("cbIRDat").ValueEx & "', "
                sUpdate &= " ""U_SOAPRTOPT""    = '" & .Item("cbSOAPrt").ValueEx & "', "
                sUpdate &= " ""U_EmailType""    = '" & .Item("cbEmailTyp").ValueEx & "', "
                sUpdate &= " ""U_PlainText""    = '" & .Item("tbNotes").ValueEx.ToString.Replace("'", "''") & "', "
                sUpdate &= " ""U_InvEmailType"" = '" & .Item("cbInvType").ValueEx & "', "
                sUpdate &= " ""U_InvPlainText"" = '" & .Item("tbInvText").ValueEx.ToString.Replace("'", "''") & "', "
                sUpdate &= " ""U_RinEmailType"" = '" & .Item("cbRinType").ValueEx & "', "
                sUpdate &= " ""U_RinPlainText"" = '" & .Item("tbRinText").ValueEx.ToString.Replace("'", "''") & "', "
                sUpdate &= " ""U_DpiEmailType"" = '" & .Item("cbDpiType").ValueEx & "', "
                sUpdate &= " ""U_DpiPlainText"" = '" & .Item("tbDpiText").ValueEx.ToString.Replace("'", "''") & "', "
                sUpdate &= " ""U_CardOption""   = '" & .Item("cbCardOp").ValueEx.ToString.Trim & "' "

                oUpdate = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oUpdate.DoQuery(sUpdate)

                ' NEW PARAMETER - U_SOAPRTOPT
                ' ==================================================================================
                Try
                    sUpdate = "  UPDATE ""@NCM_NEW_SETTING"" SET ""U_SOAPRTOPT"" = '" & .Item("cbSOAPrt").ValueEx & "' "
                    oUpdate = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oUpdate.DoQuery(sUpdate)
                Catch ex As Exception

                End Try
                ' ==================================================================================

                For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                    oMtrxBSC.GetLineData(i)
                    oUpdate = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    sUpdate = "  UPDATE ""@NCM_RPT_CONFIG"" "
                    sUpdate &= " SET    ""FILEPATH"" ='" & .Item("xFilePath").ValueEx.Replace("'", "''") & "'"
                    sUpdate &= " WHERE  ""RPTCODE""  ='" & .Item("xRptCode").ValueEx & "'"
                    oUpdate.DoQuery(sUpdate)
                Next

                For i As Integer = 1 To oMtrxADV.VisualRowCount Step 1
                    oMtrxADV.GetLineData(i)
                    oUpdate = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                    sUpdate = "  UPDATE ""@NCM_RPT_CONFIG"" "
                    sUpdate &= " SET    ""FILEPATH"" ='" & .Item("yFilePath").ValueEx.Replace("'", "''") & "'"
                    sUpdate &= " WHERE  ""RPTCODE""  ='" & .Item("yRptCode").ValueEx & "'"
                    oUpdate.DoQuery(sUpdate)
                Next
            End With

            oUpdate = Nothing
            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[UpdateReport] : " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub CopyFromPrevious()
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            Dim sGSTCurr As String = ""
            Dim sInvDet As String = "N"
            Dim sInvTaxDate As String = "Y"
            Dim sIRAInvDet As String = "N"
            Dim sIRAInvTaxDate As String = "Y"
            Dim sFilePath As String = ""

            oFormCFG.Freeze(True)
            sQuery = " SELECT * FROM ""@NCM_SETTING"" "
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                With oFormCFG.DataSources.UserDataSources

                    '' AR SOA
                    sFilePath = oRecord.Fields.Item("U_ARSOAName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AR_SOA" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AP SOA
                    sFilePath = oRecord.Fields.Item("U_APSOAName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AP_SOA" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AP Ageing Summary
                    sFilePath = oRecord.Fields.Item("U_APAGESName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AP_AGEING_SUMMARY" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AP Ageing Details
                    sFilePath = oRecord.Fields.Item("U_APAGEDName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AP_AGEING_DETAILS" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AR Ageing Summary (5 & 6 buckets)
                    sFilePath = oRecord.Fields.Item("U_ARAGESName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AR_AGEING_SUMMARY" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        If .Item("xRptCode").ValueEx = "AR_AGEING6B_SUMMARY" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AR Ageing Details (5 & 6 buckets)
                    sFilePath = oRecord.Fields.Item("U_ARAGEDName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AR_AGEING_DETAILS" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        If .Item("xRptCode").ValueEx = "AR_AGEING6B_DETAILS" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' GST Report
                    sFilePath = oRecord.Fields.Item("U_RPT_GST_Name").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "RPT_GST" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AP Payment __ U_InAPPayment
                    sFilePath = oRecord.Fields.Item("U_InAPPayment").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AP_PAYMENT" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' AR Payment __ U_InARPayment
                    sFilePath = oRecord.Fields.Item("U_InARPayment").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "AR_PAYMENT" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' Payment Voucher __ U_ReportName
                    sFilePath = oRecord.Fields.Item("U_ReportName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "PAYMENT_VOUCHER" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' Remittance Advice __ U_RemitName
                    sFilePath = oRecord.Fields.Item("U_RemitName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "REMITTANCE_ADVICE" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' Payment Voucher Range __ U_PVRangeName
                    sFilePath = oRecord.Fields.Item("U_PVRangeName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "PAYMENT_VOUCHER_RANGE" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' Draft Payment Voucher __ U_PVDraftRptName
                    sFilePath = oRecord.Fields.Item("U_PVDraftRptName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "DRAFT_PV" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' Official Receipt __ U_IRAName
                    sFilePath = oRecord.Fields.Item("U_IRAName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "OFFICIAL_RECEIPT" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' Receipt and Payment ___ U_RecpPaymName 
                    sFilePath = oRecord.Fields.Item("U_RecpPaymName").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "RECEIPT_PAYMENT" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' SAR Fifo Summary ___ U_SAR_FIFO1_Name 
                    sFilePath = oRecord.Fields.Item("U_SAR_FIFO1_Name").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "SAR_FIFO_SUMMARY" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' SAR Fifo Details ___ U_SAR_FIFO2_Name
                    sFilePath = oRecord.Fields.Item("U_SAR_FIFO2_Name").Value
                    For i As Integer = 1 To oMtrxBSC.VisualRowCount Step 1
                        oMtrxBSC.GetLineData(i)
                        If .Item("xRptCode").ValueEx = "SAR_FIFO_DETAILS" Then
                            .Item("xFilePath").ValueEx = sFilePath
                        End If
                        oMtrxBSC.SetLineData(i)
                    Next

                    '' GPA Report - Internal - No need to copy
                    '' GL Listing Report - Internal - No need to copy
                    '' AR SOA Email - No
                    '' AR SOA PRoject - No
                    '' MRP Report - No - ADV
                    '' Hydac IAR - No - ADV
                    '' Hydac RLR - No - ADV
                    '' Hydac WAR - No - ADV
                    '' SAR Moving Avg Summary - No
                    '' SAR Moving Avg Details - No
                    '' SAR Enquiry - No
                    '' SAR TM V1 - No
                    '' SAR TM V2 - No
                    '' SAR TM V3 - No

                End With
            End If

            sQuery = "  SELECT  IFNULL(""U_GSTCurr"",'') As GSTCurr, IFNULL(""U_InvDetail"",'N') AS InvDet, "
            sQuery &= "         IFNULL(""U_TaxDate"",'Y') As InvTaxDate, IFNULL(""U_IRAInvDetail"",'N') AS IRAInvDet, "
            sQuery &= "         IFNULL(""U_IRATaxDate"",'Y') As IRAInvTaxDate "
            sQuery &= " FROM    ""@NCM_SETTING"" "
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                sGSTCurr = oRecord.Fields.Item(0).Value
                sInvDet = oRecord.Fields.Item(1).Value
                sInvTaxDate = oRecord.Fields.Item(2).Value
                sIRAInvDet = oRecord.Fields.Item(3).Value
                sIRAInvTaxDate = oRecord.Fields.Item(4).Value

                With oFormCFG.DataSources.UserDataSources
                    .Item("cbGSTCurr").ValueEx = sGSTCurr
                    .Item("cbPVInv").ValueEx = sInvDet
                    .Item("cbPVDat").ValueEx = sInvTaxDate
                    .Item("cbIRInv").ValueEx = sIRAInvDet
                    .Item("cbIRDat").ValueEx = sIRAInvTaxDate
                End With
                oFormCFG.Mode = BoFormMode.fm_UPDATE_MODE
                SBO_Application.StatusBar.SetText("Previous settings from V04 has been copied to the new table.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
            End If

            oFormCFG.Freeze(False)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[CopyFromPreviousSettings] : " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
            oFormCFG.Freeze(False)
        End Try
    End Sub
    Private Sub ClearPreviousSetting()
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty

            Dim array(50) As String
            array(0) = "U_GSTCurr"
            array(1) = "U_InvDetail"
            array(2) = "U_TaxDate"
            array(3) = "U_IRAInvDetail"
            array(4) = "U_IRATaxDate"
            array(5) = "U_InAPPayment"
            array(6) = "U_InARPayment"
            array(7) = "U_InPV"
            array(8) = "U_InRA"
            array(9) = "U_InGST"

            array(10) = "U_InARAging"
            array(11) = "U_InAPAging"
            array(12) = "U_InARSOA"
            array(13) = "U_InAPSOA"
            array(14) = "U_InOMARSOA"
            array(15) = "U_InIRA"
            array(16) = "U_InARAging6B"
            array(17) = "U_InSOAEmail"
            array(18) = "U_InTMSAR"
            array(19) = "U_InPVDraft"

            array(20) = "U_InRecpPaym"
            array(21) = "U_InGPA"
            array(22) = "U_InSOAPrj"
            array(23) = "U_PrintPV"
            array(24) = "U_PrintPVDraft"
            array(25) = "U_PrintIRA"
            array(26) = "U_SAR_STRUC_NAME"
            array(27) = "U_RptSOAPrj"
            array(28) = "U_RptGPA"
            array(29) = "U_TM_SAR_Name"

            array(30) = "U_OMARSOAName"
            array(31) = "U_RPT_GST_Name"
            array(32) = "U_SAR_FIFO1_Name"
            array(33) = "U_SAR_FIFO2_Name"
            array(34) = "U_RecpPaymName"
            array(35) = "U_IRAName"
            array(36) = "U_PVDraftRptName"
            array(37) = "U_RemitName"
            array(38) = "U_PVRangeName"
            array(39) = "U_ReportName"

            array(40) = "U_ARSOAName"
            array(41) = "U_APSOAName"
            array(42) = "U_APAGESName"
            array(43) = "U_APAGEDName"
            array(44) = "U_ARAGESName"
            array(45) = "U_ARAGEDName"
            array(46) = ""
            array(47) = ""
            array(48) = ""
            array(49) = ""


            For Each sValue As String In array
                If sValue <> "" Then
                    sQuery = "  IF Exists(select * from sys.columns where ""Name"" = N'" & sValue & "' and Object_ID = Object_ID(N'""@NCM_SETTING""'))"
                    sQuery &= " BEGIN "
                    sQuery &= " ALTER TABLE ""@NCM_SETTING"" DROP COLUMN """ & sValue & """ "
                    sQuery &= " END "

                    oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(sQuery)
                End If
            Next

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ClearPreviousSetting] : " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Event Handlers"
    Friend Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.EventType = BoEventTypes.et_FORM_RESIZE Then
                Try
                    oFormCFG = SBO_Application.Forms.Item(FRM_RPT_CONFIG)
                    oMtrxBSC = oFormCFG.Items.Item("mxBSC").Specific
                    oMtrxADV = oFormCFG.Items.Item("mxADV").Specific

                    oMtrxBSC.AutoResizeColumns()
                    oMtrxADV.AutoResizeColumns()
                Catch ex As Exception

                End Try
            End If

            If pVal.Before_Action = True Then
                If pVal.ItemUID = "1" Then
                    If pVal.EventType = BoEventTypes.et_ITEM_PRESSED Then
                        If pVal.FormMode = BoFormMode.fm_UPDATE_MODE Then
                            BubbleEvent = False
                            If UpdateReport() Then
                                oFormCFG.Mode = BoFormMode.fm_OK_MODE
                                SBO_Application.StatusBar.SetText("Operation is successfully completed.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success)
                            End If
                        End If
                    End If
                End If
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "flOne"
                                oFormCFG.PaneLevel = 1
                                g_iPaneLvl = 1
                                oMtrxBSC.AutoResizeColumns()

                            Case "flTwo"
                                oFormCFG.PaneLevel = 2
                                g_iPaneLvl = 2
                                oMtrxADV.AutoResizeColumns()

                            Case "flThree"
                                oFormCFG.PaneLevel = 3
                                g_iPaneLvl = 3

                            Case "flFour"
                                oFormCFG.PaneLevel = 4
                                g_iPaneLvl = 4

                        End Select
                    Case BoEventTypes.et_ITEM_PRESSED
                        Select Case pVal.ItemUID
                            Case "btCopy"
                                CopyFromPrevious()
                            Case "btClear"
                                Dim iReturn As Integer = 0
                                iReturn = SBO_Application.MessageBox("Please confirm if you want to clear the previous setting permanently.", 2, "&Yes", "&No")
                                If iReturn = 1 Then
                                    ClearPreviousSetting()
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            BubbleEvent = False
            SBO_Application.StatusBar.SetText("[ItemEvent] : " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
        Return BubbleEvent
    End Function


#End Region

End Class

