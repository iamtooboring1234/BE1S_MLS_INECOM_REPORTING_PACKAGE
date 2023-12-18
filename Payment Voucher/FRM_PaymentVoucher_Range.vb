Imports System.IO
Imports System.Threading
Imports System.Data.Common
Imports System.Globalization
Imports System.xml

Public Class FRM_PaymentVoucher_Range

#Region "Global Variables"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private sErrMsg As String
    Private lErrCode As Integer
    Private g_sReportFilename As String
    Private g_bIsShared As Boolean = False

    Private g_sReportFilename_Email As String = ""
    Private g_bIsShared_Email As Boolean = False

    Private dsPAYMENT As DataSet
    Private g_sDocNum As String = ""
    Private g_sDocEntry As String = ""
    Private g_sSeries As String = ""
    Private g_sDocType As String = ""
    Private AsAtDate As DateTime

    Private g_StructureFilename As String = ""
    Private myThread As System.Threading.Thread
    Private Const ITEM01 As String = "txtSDocNum"
    Private Const ITEM02 As String = "txtEDocNum"
    Private Const ITEM03 As String = "txtSBPCode"
    Private Const ITEM04 As String = "txtEBPCode"
    Private Const ITEM05 As String = "txtSDate"
    Private Const ITEM06 As String = "txtEDate"
    Private Const ITEM07 As String = "cboType"
    Private Const ITEM08 As String = "btnPrint"
    Private Const ITEM09 As String = "2"
    Private Const ITEM10 As String = "cflSBPCode"
    Private Const ITEM11 As String = "cflEBPCode"
    Private Const ITEM12 As String = "chkCancel"
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox

    Private g_bShowDetails As Boolean = False
    Private g_sShowTaxDate As String = ""
#End Region

#Region "Constructors"
    Public Sub New()
        Me.SBO_Application = SubMain.SBO_Application
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region

#Region "Populate Data"
    Friend Sub LoadForm()
        Try
            If LoadFromXML("Inecom_SDK_Reporting_Package.SRF_NCM_PAYCHER.srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading Form...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oForm = SBO_Application.Forms.Item(FRM_NCM_PV_RANGE)
                oForm.EnableMenu(MenuID.Add, False)
                oForm.EnableMenu(MenuID.Find, False)
                oForm.EnableMenu(MenuID.Remove_Record, True)
                oForm.EnableMenu(MenuID.Find, True)
                oForm.EnableMenu(MenuID.Add, True)
                oForm.EnableMenu(MenuID.Paste, True)
                oForm.EnableMenu(MenuID.Copy, True)
                oForm.EnableMenu(MenuID.Cut, True)

                oForm.Items.Item("20").Visible = True
                oForm.Items.Item("cbLayout").Visible = True

                InitializeItem()

                oForm.Items.Item("lkEntFr").Visible = True
                oForm.Items.Item("lkEntTo").Visible = True

                oForm.Items.Item("txtSDocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                SBO_Application.StatusBar.SetText(String.Empty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                oForm.Visible = True
            Else
                ' Loading .srf file failed most likely it is because the form is already opened
                Try
                    oForm = SBO_Application.Forms.Item(FRM_NCM_PV_RANGE)
                    If oForm.Visible Then
                        oForm.Select()
                    Else
                        oForm.Close()
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadForm] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub SetupChooseFromList(ByVal sInputDocumentType As String)
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim sDocType As String = "C"


        Try
            Select Case sInputDocumentType
                Case "C"

                    oEditLn = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).Specific
                    oEditLn.ChooseFromListUID = "CFL_CUSTFR"
                    oEditLn.ChooseFromListAlias = "CardCode"

                    oEditLn = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode)).Specific
                    oEditLn.ChooseFromListUID = "CFL_CUSTTO"
                    oEditLn.ChooseFromListAlias = "CardCode"

                Case "S"

                    oEditLn = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).Specific
                    oEditLn.ChooseFromListUID = "CFL_SUPPFR"
                    oEditLn.ChooseFromListAlias = "CardCode"

                    oEditLn = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode)).Specific
                    oEditLn.ChooseFromListUID = "CFL_SUPPTO"
                    oEditLn.ChooseFromListAlias = "CardCode"
                Case Else

                    oEditLn = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).Specific
                    oEditLn.ChooseFromListUID = GetItemUID(PaymentVoucherRangeItems.CFL_StartingBPCode)
                    oEditLn.ChooseFromListAlias = "CardCode"

                    oEditLn = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode)).Specific
                    oEditLn.ChooseFromListUID = GetItemUID(PaymentVoucherRangeItems.CFL_EndingBPCode)
                    oEditLn.ChooseFromListAlias = "CardCode"
            End Select

        Catch ex As Exception
            Throw New Exception("[RPV].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub

    Friend Sub InitializeItem()
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oLink As SAPbouiCOM.LinkedButton
        Dim oCbox As SAPbouiCOM.ComboBox
        Dim oChck As SAPbouiCOM.CheckBox
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition

        oCFLs = oForm.ChooseFromLists

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = "2"
        oCFLCreation.UniqueID = GetItemUID(PaymentVoucherRangeItems.CFL_StartingBPCode)
        oCFL = oCFLs.Add(oCFLCreation)

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = "2"
        oCFLCreation.UniqueID = GetItemUID(PaymentVoucherRangeItems.CFL_EndingBPCode)
        oCFL = oCFLs.Add(oCFLCreation)

        ' ========================================================================

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = "2"
        oCFLCreation.UniqueID = "CFL_CUSTFR"
        oCFL = oCFLs.Add(oCFLCreation)

        oCons = New SAPbouiCOM.Conditions
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "C"
        oCFL.SetConditions(oCons)

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = "2"
        oCFLCreation.UniqueID = "CFL_CUSTTO"
        oCFL = oCFLs.Add(oCFLCreation)

        oCons = New SAPbouiCOM.Conditions
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "C"
        oCFL.SetConditions(oCons)

        ' ========================================================================

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = "2"
        oCFLCreation.UniqueID = "CFL_SUPPFR"
        oCFL = oCFLs.Add(oCFLCreation)

        oCons = New SAPbouiCOM.Conditions
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "S"
        oCFL.SetConditions(oCons)

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = "2"
        oCFLCreation.UniqueID = "CFL_SUPPTO"
        oCFL = oCFLs.Add(oCFLCreation)

        oCons = New SAPbouiCOM.Conditions
        oCon = oCons.Add()
        oCon.Alias = "CardType"
        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
        oCon.CondVal = "S"
        oCFL.SetConditions(oCons)

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = 10
        oCFLCreation.UniqueID = "CFL_BGFR"
        oCFL = oCFLs.Add(oCFLCreation)

        oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
        oCFLCreation.MultiSelection = False
        oCFLCreation.ObjectType = 10
        oCFLCreation.UniqueID = "CFL_BGTO"
        oCFL = oCFLs.Add(oCFLCreation)

        With oForm.DataSources.UserDataSources
            .Add("txtSDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add("txtEDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add("txtSBPGrp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add("txtEBPGrp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add("tbEntFr", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            .Add("tbEntTo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            .Add("cbWizard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("ckWizard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            .Add("cbLayout", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2)

            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate), SAPbouiCOM.BoDataType.dt_DATE, 254)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate), SAPbouiCOM.BoDataType.dt_DATE, 254)
            .Add(GetItemUID(PaymentVoucherRangeItems.CBO_DocType), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            .Add(GetItemUID(PaymentVoucherRangeItems.CHK_IncludeCancel), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
        End With

        oChck = oForm.Items.Item("ckWizard").Specific
        oChck.DataBind.SetBound(True, String.Empty, "ckWizard")
        oChck.ValOff = "0"
        oChck.ValOn = "1"
        oForm.DataSources.UserDataSources.Item("ckWizard").ValueEx = "0"

        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sRec As String = ""
        Dim sFirstWizard As String = ""
        sRec = "SELECT ""IdNumber"", ""WizardName"" FROM ""OPWZ"" WHERE ""OutgoType"" = 'Y' ORDER BY ""IdNumber"" "
        oRec.DoQuery(sRec)

        oCbox = oForm.Items.Item("cbWizard").Specific
        oCbox.DataBind.SetBound(True, String.Empty, "cbWizard")
        If oRec.RecordCount > 0 Then
            oRec.MoveFirst()
            sFirstWizard = oRec.Fields.Item(0).Value.ToString.Trim
            While Not oRec.EoF
                oCbox.ValidValues.Add(oRec.Fields.Item(0).Value.ToString.Trim, oRec.Fields.Item(1).Value.ToString.Trim)
                oRec.MoveNext()
            End While
            oCbox.Select(sFirstWizard, SAPbouiCOM.BoSearchKey.psk_ByValue)
        End If
        oRec = Nothing

        If sFirstWizard.Length > 0 Then
            oForm.DataSources.UserDataSources.Item("cbWizard").ValueEx = sFirstWizard
        End If

        oEdit = oForm.Items.Item("tbEntFr").Specific
        oEdit.DataBind.SetBound(True, "", "tbEntFr")
        oEdit = oForm.Items.Item("tbEntTo").Specific
        oEdit.DataBind.SetBound(True, "", "tbEntTo")

        oLink = oForm.Items.Item("lkEntFr").Specific
        oLink.LinkedObject = 46
        oLink = oForm.Items.Item("lkEntTo").Specific
        oLink.LinkedObject = 46

        oForm.Items.Item("lkEntFr").LinkTo = "tbEntFr"
        oForm.Items.Item("lkEntTo").LinkTo = "tbEntTo"
        ' =============================================================
        oEdit = oForm.Items.Item("txtSDocNum").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtSDocNum")
        oEdit = oForm.Items.Item("txtEDocNum").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtEDocNum")
        ' =============================================================
        oEdit = oForm.Items.Item("txtSBPGrp").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtSBPGrp")
        oEdit.ChooseFromListUID = "CFL_BGFR"
        oEdit.ChooseFromListAlias = "GroupCode"
        ' =============================================================
        oEdit = oForm.Items.Item("txtEBPGrp").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtEBPGrp")
        oEdit.ChooseFromListUID = "CFL_BGTO"
        oEdit.ChooseFromListAlias = "GroupCode"
        ' =============================================================

        oEdit = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).Specific
        oEdit.DataBind.SetBound(True, String.Empty, GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode))
        oEdit.ChooseFromListUID = GetItemUID(PaymentVoucherRangeItems.CFL_StartingBPCode)
        oEdit.ChooseFromListAlias = "CardCode"

        oEdit = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode)).Specific
        oEdit.DataBind.SetBound(True, String.Empty, GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode))
        oEdit.ChooseFromListUID = GetItemUID(PaymentVoucherRangeItems.CFL_EndingBPCode)
        oEdit.ChooseFromListAlias = "CardCode"

        oEdit = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate)).Specific
        oEdit.DataBind.SetBound(True, String.Empty, GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate))
        oEdit = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate)).Specific
        oEdit.DataBind.SetBound(True, String.Empty, GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate))

        oCombo = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.CBO_DocType)).Specific
        oCombo.DataBind.SetBound(True, String.Empty, GetItemUID(PaymentVoucherRangeItems.CBO_DocType))
        oCombo.ValidValues.Add(GetDocType(PaymentVoucherRangeDocTypes.All), "ALL")
        oCombo.ValidValues.Add(GetDocType(PaymentVoucherRangeDocTypes.Account), "Account")
        oCombo.ValidValues.Add(GetDocType(PaymentVoucherRangeDocTypes.Customer), "Customer")
        oCombo.ValidValues.Add(GetDocType(PaymentVoucherRangeDocTypes.Supplier), "Supplier")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

        oCombo = oForm.Items.Item("cbLayout").Specific
        oCombo.DataBind.SetBound(True, "", "cbLayout")
        oCombo.ValidValues.Add("PV", "Payment Voucher")
        oCombo.ValidValues.Add("RA", "Remittance Advice")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx = "PV"

        oCheck = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.CHK_IncludeCancel)).Specific
        oCheck.DataBind.SetBound(True, String.Empty, GetItemUID(PaymentVoucherRangeItems.CHK_IncludeCancel))
        oCheck.ValOff = "0"
        oCheck.ValOn = "1"

        oForm.DataSources.UserDataSources.Item("tbEntFr").ValueEx = 0
        oForm.DataSources.UserDataSources.Item("tbEntTo").ValueEx = 0

    End Sub
    Private Function GetItemUID(ByVal PaymentVoucherRangeItem As PaymentVoucherRangeItems) As String
        Select Case PaymentVoucherRangeItem
            Case PaymentVoucherRangeItems.TXT_StartingDocNum
                Return ITEM01
            Case PaymentVoucherRangeItems.TXT_EndingDocNum
                Return ITEM02
            Case PaymentVoucherRangeItems.TXT_StartingBPCode
                Return ITEM03
            Case PaymentVoucherRangeItems.TXT_EndingBPCode
                Return ITEM04
            Case PaymentVoucherRangeItems.TXT_StartingDate
                Return ITEM05
            Case PaymentVoucherRangeItems.TXT_EndingDate
                Return ITEM06
            Case PaymentVoucherRangeItems.CBO_DocType
                Return ITEM07
            Case PaymentVoucherRangeItems.BTN_PRINT
                Return ITEM08
            Case PaymentVoucherRangeItems.BTN_CANCEL
                Return ITEM09
            Case PaymentVoucherRangeItems.CFL_StartingBPCode
                Return ITEM10
            Case PaymentVoucherRangeItems.CFL_EndingBPCode
                Return ITEM11
            Case PaymentVoucherRangeItems.CHK_IncludeCancel
                Return ITEM12
        End Select
        Return String.Empty
    End Function
    Private Function GetDocType(ByVal PaymentVoucherRangeDocType As PaymentVoucherRangeDocTypes) As String
        Dim i As Integer = PaymentVoucherRangeDocType
        Return i.ToString()
    End Function
    Private Function IsParametersValid() As Boolean
        Try
            oForm.ActiveItem = GetItemUID(PaymentVoucherRangeItems.TXT_StartingDocNum)
            Dim oRecordsetLn As SAPbobsCOM.Recordset = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty
            Dim sQuery As String = String.Empty

            sStart = oForm.DataSources.UserDataSources.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDocNum)).ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDocNum)).ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Doc Num from is greater than Doc Num to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = GetItemUID(PaymentVoucherRangeItems.TXT_StartingDocNum)
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).ValueEx

            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("BP Code from is greater than BP Code to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate)).ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate)).ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Doc Date from is greater than Doc Date to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate)
                    Return False
                End If
            End If
            SBO_Application.StatusBar.SetText(String.Empty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[PaymentVoucherRange].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
        Return False
    End Function
#End Region

#Region "Print Report"

    Private Function IsSharedFileExist(ByVal sAction As String) As Boolean
        Try
            Dim sLayoutType As String = oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx
            g_sReportFilename = ""
            g_StructureFilename = ""

            Select Case sLayoutType
                Case "PV"
                    Select Case sAction
                        Case "Preview"
                            g_sReportFilename = GetSharedFilePath(ReportName.PV_Range)
                        Case "Email"
                            g_sReportFilename = GetSharedFilePath(ReportName.PV)
                    End Select

                Case "RA"
                    Select Case sAction
                        Case "Preview"
                            g_sReportFilename = GetSharedFilePath(ReportName.RA_Range)
                        Case "Email"
                            g_sReportFilename = GetSharedFilePath(ReportName.RA)
                    End Select

            End Select
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If

            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[RPV.GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub Print_Report()

        oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.BTN_PRINT)).Enabled = False
        Dim sFinalExportPath As String = ""
        Dim sFinalFileName As String = ""
        Dim iCount As Integer = 0
        Dim bIsWizard As Boolean = False
        Dim sWizardCode As String = ""

        If oForm.DataSources.UserDataSources.Item("ckWizard").ValueEx = "1" Then
            bIsWizard = True
            sWizardCode = oForm.DataSources.UserDataSources.Item("cbWizard").ValueEx.ToString.Trim
        End If

        Try
            Dim frm As Hydac_FormViewer = New Hydac_FormViewer
            Dim bIsContinue As Boolean = False
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\RPV_{1}.pdf"
            Dim sPathFormatSingle As String = "{0}\RPV_{1}_{2}_{3}.pdf"
            Dim sCurrDate As String = ""
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
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\RPV\" & oCompany.CompanyDB

            If oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx.Trim = "RA" Then
                sPathFormat = "{0}\RRA_{1}.pdf"
                sPathFormatSingle = "{0}\RRA_{1}_{2}_{3}.pdf"
                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\RRA\" & oCompany.CompanyDB
            End If

            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\RPV_" & sCurrDate & "_" & sCurrTime & ".pdf"

            If oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx.Trim = "RA" Then
                sFinalFileName = di.FullName & "\RRA_" & sCurrDate & "_" & sCurrTime & ".pdf"
            End If
            ' ===============================================================================

            Try

                Dim sTemp As String = String.Empty
                Dim iTemp As Integer = 0
                Dim sDocNumS As String = String.Empty
                Dim sDocNumE As String = String.Empty
                Dim sBPCodeS As String = String.Empty
                Dim sBPCodeE As String = String.Empty
                Dim sDocDateS As String = String.Empty
                Dim sDocDateE As String = String.Empty
                Dim dtDocDateS As DateTime
                Dim dtDocDateE As DateTime
                Dim iIsIncludeCancel As Integer = 0
                Dim myPaymentVoucherDocType As PaymentVoucherRangeDocTypes = PaymentVoucherRangeDocTypes.Supplier

                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery("SELECT ""U_INVDETAIL"", ""U_TAXDATE"" FROM """ & oCompany.CompanyDB & """.""@NCM_NEW_SETTING""")
                g_bShowDetails = IIf(oRec.Fields.Item(0).Value = "Y", True, False)
                g_sShowTaxDate = oRec.Fields.Item(1).Value

                With oForm.DataSources.UserDataSources
                    sDocNumS = .Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDocNum)).ValueEx
                    sDocNumE = .Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDocNum)).ValueEx
                    sBPCodeS = .Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode)).ValueEx
                    sBPCodeE = .Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode)).ValueEx
                    sDocDateS = DirectCast(oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate)).Specific, SAPbouiCOM.EditText).Value
                    If (sDocDateS.Length > 0) Then
                        dtDocDateS = DateTime.ParseExact(.Item(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate)).ValueEx, "yyyyMMdd", Nothing)
                    End If

                    sDocDateE = DirectCast(oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate)).Specific, SAPbouiCOM.EditText).Value
                    If (sDocDateE.Length > 0) Then
                        dtDocDateE = DateTime.ParseExact(.Item(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate)).ValueEx, "yyyyMMdd", Nothing)
                    End If

                    If (DirectCast(oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.CHK_IncludeCancel)).Specific, SAPbouiCOM.CheckBox).Checked) Then
                        iIsIncludeCancel = 1
                    End If
                    sTemp = .Item(GetItemUID(PaymentVoucherRangeItems.CBO_DocType)).ValueEx
                    iTemp = Integer.Parse(sTemp)
                End With

                Dim oLoop As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim sLoop As String = ""

                Select Case bIsWizard
                    Case True
                        sLoop = "  SELECT   T1.""DocEntry"", T1.""DocNum"", T1.""DocType"", T1.""Series"", T1.""CardCode"", T2.""CardName"", "
                        sLoop &= "          T1.""DocCurr"", IFNULL(T2.""U_PV_MailTo"",'') ""EmailTo"", T0.""PymAmount"" ""DocTotal"", T0.""PymNum"", "
                        sLoop &= "          IFNULL(T1.""U_AcctMailTo"",'') ""AcctEmailTo"" "
                        sLoop &= " FROM """ & oCompany.CompanyDB & """.""PWZ4"" T0 "
                        sLoop &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OVPM"" T1 On T0.""RctId""     = T1.""DocNum"" "
                        sLoop &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T2 On T1.""CardCode""  = T2.""CardCode""  "
                        sLoop &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T3 On T2.""GroupCode"" = T3.""GroupCode"" "
                        sLoop &= " WHERE 1=1 "
                        sLoop &= " And T0.""IdEntry"" = '" & sWizardCode & "' "
                        sLoop &= " AND T0.""ObjType"" = 46 "
                        sLoop &= " ORDER BY T0.""PymNum"" "

                    Case False
                        sLoop = "  SELECT   T1.""DocEntry"", T1.""DocNum"", T1.""DocType"", T1.""Series"", T1.""CardCode"", T2.""CardName"", "
                        sLoop &= "          T1.""DocCurr"", IFNULL(T2.""U_PV_MailTo"",'') ""EmailTo"", "
                        sLoop &= "          IFNULL(T1.""U_AcctMailTo"",'') ""AcctEmailTo"", "
                        sLoop &= "          SUM(T1.""DocTotal"") ""DocTotal""  "
                        sLoop &= " FROM """ & oCompany.CompanyDB & """.""OVPM"" T1 "
                        sLoop &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T2 ON T1.""CardCode""  = T2.""CardCode""  "
                        sLoop &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRG"" T3 ON T2.""GroupCode"" = T3.""GroupCode"" "
                        sLoop &= " WHERE 1=1 "

                        If iIsIncludeCancel = 1 Then
                            ' sLoop &= " And ""Canceled"" = 'Y' " ' doesnt matter Y or N
                        Else
                            sLoop &= " AND T1.""Canceled"" = 'N'  "
                        End If

                        oEdit = oForm.Items.Item("txtSBPGrp").Specific
                        If oEdit.Value.ToString.Trim() <> "" Then
                            sQuery &= " AND T3.""GroupCode"" >= '" & oEdit.Value.ToString.Trim & "' "
                        End If

                        oEdit = oForm.Items.Item("txtEBPGrp").Specific
                        If oEdit.Value.ToString().Trim() <> "" Then
                            sQuery &= " AND T3.""GroupCode"" <= '" & oEdit.Value.ToString.Trim & "' "
                        End If

                        Select Case sTemp
                            Case "1"
                                sLoop &= " AND T1.""DocType"" = 'A' "
                            Case "2"
                                sLoop &= " AND T1.""DocType"" = 'C' "
                            Case "3"
                                sLoop &= " AND T1.""DocType"" = 'S' "
                        End Select

                        If sBPCodeS.Trim.Length > 0 Then
                            sLoop &= " AND T1.""CardCode"" >= '" & sBPCodeS.Trim & "' "
                        End If
                        If sBPCodeE.Trim.Length > 0 Then
                            sLoop &= " AND T1.""CardCode"" <= '" & sBPCodeE.Trim & "' "
                        End If
                        If sDocNumS.Trim.Length > 0 Then
                            sLoop &= " AND T1.""DocNum"" >= '" & sDocNumS.Trim & "' "
                        End If
                        If sDocNumE.Trim.Length > 0 Then
                            sLoop &= " AND T1.""DocNum"" <= '" & sDocNumE.Trim & "' "
                        End If
                        If sDocDateS.Trim.Length > 0 Then
                            sLoop &= " AND T1.""DocDate"" >= '" & sDocDateS.Trim & "' "
                        End If
                        If sDocDateE.Trim.Length > 0 Then
                            sLoop &= " AND T1.""DocDate"" <= '" & sDocDateE.Trim & "' "
                        End If

                        sLoop &= " GROUP BY T1.""DocEntry"", T1.""DocNum"", T1.""DocType"", T1.""Series"", T1.""CardCode"", T2.""CardName"", T1.""DocCurr"", T2.""U_PV_MailTo"", T1.""U_AcctMailTo""  "
                        sLoop &= " ORDER BY T1.""DocEntry"", T1.""DocNum"", T1.""DocType"", T1.""Series"", T1.""CardCode"", T2.""CardName"", T1.""DocCurr"", T2.""U_PV_MailTo"", T1.""U_AcctMailTo""  "
                End Select

                iCount = 2

                If (IsIncludeModule(ReportName.PV_Mass_Email)) Then
                    iCount = SBO_Application.MessageBox("Please select your option." & vbNewLine & "1. Click ""Yes"" to send email." & vbNewLine & "2. Click ""No"" to preview only.", 1, "Yes", "No", String.Empty)
                End If

                Select Case iCount
                    Case 1
                        g_bIsShared = IsSharedFileExist("Email")
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

                        ' EMAIL
                        oLoop = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oLoop.DoQuery(sLoop)

                        If oLoop.RecordCount > 0 Then
                            SBO_Application.StatusBar.SetText("Preparing the email, please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            Dim ds As New dsPVEmail()

                            While Not oLoop.EoF
                                g_sDocEntry = oLoop.Fields.Item("DocEntry").Value.ToString()
                                g_sDocNum = oLoop.Fields.Item("DocNum").Value.ToString()
                                g_sSeries = oLoop.Fields.Item("Series").Value.ToString()
                                g_sDocType = oLoop.Fields.Item("DocType").Value.ToString()

                                SBO_Application.StatusBar.SetText("Generating " & oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx.Trim & " #" & g_sDocNum & "...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                If PrepareDatasetSingle() Then
                                    With frm
                                        .LayoutType = oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx.Trim

                                        Select Case oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx.Trim
                                            Case "PV"
                                                .ReportName = ReportName.PV
                                                .ReportNamePV = g_sReportFilename
                                            Case "RA"
                                                .ReportName = ReportName.RA
                                                .ReportNameRA = g_sReportFilename
                                        End Select

                                        .IsShared = g_bIsShared
                                        .ExportPath = sFinalFileName
                                        .Dataset = dsPAYMENT
                                        .DocNum = g_sDocNum
                                        .DocEntry = g_sDocEntry
                                        .Series = g_sSeries
                                        .DBUsernameViewer = DBUsername
                                        .DBPasswordViewer = DBPassword
                                        .ShowDetails = g_bShowDetails
                                        .ShowTaxDate = g_sShowTaxDate
                                        .DatabaseServer = oCompany.Server
                                        .DatabaseName = oCompany.CompanyDB
                                        .IsIncludeCancel = iIsIncludeCancel
                                        .IsExport = True
                                        .CrystalReportExportType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                                        .CrystalReportExportPath = String.Format(sPathFormatSingle, di.FullName, oLoop.Fields.Item("CardCode").Value.ToString(), System.DateTime.Now.Date.ToString("ddMMyyyy"), g_sDocNum)
                                        frm.OPEN_HANADS_PV_EMAIL()

                                    End With

                                    Dim dr As dsPVEmail.PreviewDTRow
                                    dr = ds.PreviewDT.NewPreviewDTRow()
                                    dr.Attachment = String.Format(sPathFormatSingle, di.FullName, oLoop.Fields.Item("CardCode").Value, System.DateTime.Now.Date.ToString("ddMMyyyy"), g_sDocNum)
                                    dr.Balance = oLoop.Fields.Item("DocTotal").Value
                                    dr.CardCode = oLoop.Fields.Item("CardCode").Value
                                    dr.CardName = oLoop.Fields.Item("CardName").Value
                                    dr.Currency = oLoop.Fields.Item("DocCurr").Value

                                    Select Case g_sDocType
                                        Case "A"
                                            dr.EmailTo = oLoop.Fields.Item("AcctEmailTo").Value.ToString.Trim
                                        Case Else
                                            dr.EmailTo = oLoop.Fields.Item("EmailTo").Value.ToString.Trim
                                    End Select

                                    dr.DocEntry = oLoop.Fields.Item("DocEntry").Value
                                    dr.DocNum = oLoop.Fields.Item("DocNum").Value
                                    dr.IsEmail = IIf(dr.Balance > 0, 1, 0)
                                    dr.Table.Rows.Add(dr)

                                End If

                                oLoop.MoveNext()
                            End While

                            SBO_Application.StatusBar.SetText("Showing email list, please wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            If ds.Tables(0).Rows.Count > 0 Then
                                SubMain.oFrmPVSendEmail.ReportName = ReportName.PV
                                SubMain.oFrmPVSendEmail.StatementAsAtDate = GetDateObject(GetCurrentDate)
                                SubMain.oFrmPVSendEmail.StatementDataTable = ds.PreviewDT
                                SubMain.oFrmPVSendEmail.LoadForm()
                                Hydac_FormViewer.Close()
                            End If

                        Else
                            SBO_Application.StatusBar.SetText("No data found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If

                    Case Else
                        ' PREVIEW
                        ' No need HTML 
                        ' No need email - only change Title

                        g_bIsShared = IsSharedFileExist("Preview")
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

                        SBO_Application.StatusBar.SetText("Processing Data...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oLoop.DoQuery(sLoop)
                        If oLoop.RecordCount > 0 Then
                            Dim sListDocNum As String = "("
                            Dim sListDocEntry As String = "("
                            oLoop.MoveFirst()

                            While Not oLoop.EoF
                                sListDocNum &= "'" & oLoop.Fields.Item("DocNum").Value.ToString.Trim & "',"
                                sListDocEntry &= "'" & oLoop.Fields.Item("DocEntry").Value.ToString.Trim & "',"

                                oLoop.MoveNext()
                            End While

                            sListDocNum = sListDocNum.Remove(sListDocNum.Length - 1, 1)
                            sListDocNum = sListDocNum & ")"

                            sListDocEntry = sListDocEntry.Remove(sListDocEntry.Length - 1, 1)
                            sListDocEntry = sListDocEntry & ")"
                            '=========================================
                            myPaymentVoucherDocType = iTemp

                            SBO_Application.StatusBar.SetText("Preparing Dataset...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            If PrepareDataset(sListDocNum, sListDocEntry) Then
                                With frm
                                    Select Case SBO_Application.ClientType
                                        Case SAPbouiCOM.BoClientType.ct_Desktop
                                            .ClientType = "D"
                                        Case SAPbouiCOM.BoClientType.ct_Browser
                                            .ClientType = "S"
                                    End Select

                                    SBO_Application.StatusBar.SetText("Viewing.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                    .ExportPath = sFinalFileName
                                    .Dataset = dsPAYMENT

                                    Select Case oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx.Trim
                                        Case "PV"
                                            .Text = "Payment Voucher Range Report"
                                            .ReportName = ReportName.PV_Range
                                            .ReportNamePV = g_sReportFilename
                                        Case "RA"
                                            .Text = "Remittance Advice Range Report"
                                            .ReportName = ReportName.RA_Range
                                            .ReportNameRA = g_sReportFilename
                                    End Select

                                    SBO_Application.StatusBar.SetText("Viewing..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    .LayoutType = oForm.DataSources.UserDataSources.Item("cbLayout").ValueEx
                                    .ReportName = ReportName.PV_Range
                                    .DBUsernameViewer = DBUsername
                                    .DBPasswordViewer = DBPassword
                                    .ShowDetails = g_bShowDetails
                                    .ShowTaxDate = g_sShowTaxDate
                                    .IsShared = g_bIsShared
                                    .ReportNamePV = g_sReportFilename
                                    .DocNumStart = sDocNumS
                                    .DocNumEnd = sDocNumE
                                    .BPCodeStart = sBPCodeS
                                    .BPCodeEnd = sBPCodeE
                                    .DocDateSStart = sDocDateS
                                    .DocDateSEnd = sDocDateE
                                    .DocDateStart = dtDocDateS
                                    .DocDateEnd = dtDocDateE
                                    .IsIncludeCancel = iIsIncludeCancel
                                    .PVRangeDocType = myPaymentVoucherDocType
                                    SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                                End With
                                bIsContinue = True
                            End If
                        Else
                            SBO_Application.StatusBar.SetText("No data found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        End If
                End Select


            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.BTN_PRINT)).Enabled = True
            End Try

            If iCount <> "1" Then
                If bIsContinue Then
                    SBO_Application.StatusBar.SetText("Viewing Crystal Report...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            frm.ShowDialog()

                        Case SAPbouiCOM.BoClientType.ct_Browser
                            frm.OPEN_HANADS_PV_RANGE()

                            If File.Exists(sFinalFileName) Then
                                SBO_Application.SendFileToBrowser(sFinalFileName)
                            End If
                    End Select
                End If
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[Print_Report] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Private Function PrepareDatasetSingle() As Boolean
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
            sQuery = "  SELECT T1.*, IFNULL(T2.""CardName"",'') AS ""OrigCardName"", T2.""BankCode"", T3.""BankName"", T4.""INTERNAL_K"", T4.""U_NAME"" "
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
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"", ""StreetF"", ""BlockF"", ""ZipCodeF"", ""BuildingF""  FROM """ & oCompany.CompanyDB & """.""ADM1""  "
            dtADM1 = dsPAYMENT.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)

            '--------------------------------------------------------
            sQuery = "SELECT ""FaxF"",""Phone1F"", ""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"", ""DdctOffice"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
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
            SBO_Application.StatusBar.SetText("[PrepareDatasetSingle] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function PrepareDataset(ByVal sListDocNum As String, ByVal sListDocEntry As String) As Boolean
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
            Dim dtIMAGE As System.Data.DataTable

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()

            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand

            '--------------------------------------------------------
            sQuery = " SELECT '1' ""FLAG"", '1' ""SRNO"" FROM DUMMY "
            dtIMAGE = dsPAYMENT.Tables("@NCM_IMAGE")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtIMAGE)
            '--------------------------------------------------------
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"", ""StreetF"", ""BlockF"", ""ZipCodeF"", ""BuildingF"" FROM """ & oCompany.CompanyDB & """.""ADM1""  "
            dtADM1 = dsPAYMENT.Tables("ADM1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtADM1)
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
            sQuery = " SELECT  T1.""Address"", T1.""CardCode"", T1.""CashAcct"", T1.""CashSum"", T1.""CashSumFC"", T1.""Comments"", "
            sQuery &= "     T1.""CounterRef"", T1.""NoDocSum"", T1.""NoDocSumFC"", T1.""DocCurr"", T1.""DocDate"", T1.""DocDueDate"", "
            sQuery &= "     T1.""DocEntry"", T1.""DocNum"", T1.""DocRate"", T1.""DocTotal"", T1.""DocTotalFC"", T1.""DocType"", "
            sQuery &= "     T1.""Ref1"", T1.""Ref2"", T1.""Series"", T1.""SeriesStr"", T1.""TaxDate"", T1.""TransId"", T1.""TrsfrAcct"", "
            sQuery &= "     T1.""TrsfrDate"", T1.""TrsfrRef"", T1.""TrsfrSum"", T1.""TrsfrSumFC"", T1.""LogInstanc"", T1.""DiffCurr"", "
            sQuery &= "     T1.""PrjCode"", T1.""JrnlMemo"", IFNULL(T5.""Name"",'') AS ""ContactPerson"", T1.""PayToCode"", T1.""UserSign"", "
            sQuery &= "     T1.""BcgSum"", T1.""BcgSumFC"", "
            sQuery &= "     T1.""CardName"", IFNULL(T2.""CardName"",'') AS ""OrigCardName"", T2.""BankCode"", T3.""BankName"", T4.""INTERNAL_K"", T4.""U_NAME"" "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""OVPM"" T1 "
            sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCRD"" T2 On T1.""CardCode"" = T2.""CardCode"" "
            sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""ODSC"" T3 On T2.""BankCode"" = T3.""BankCode"" "
            sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OUSR"" T4 On T1.""UserSign"" = T4.""INTERNAL_K"" "
            sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OCPR"" T5 On T1.""CardCode"" = T5.""CardCode"" AND T1.""CntctCode"" = T5.""CntctCode"" "
            sQuery &= " WHERE T1.""DocEntry"" In " & sListDocEntry & " "

            dtOPDF = dsPAYMENT.Tables("OPDF")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOPDF)

            If dtOPDF.Rows.Count > 0 Then
                SBO_Application.StatusBar.SetText("OVPM Records found : " & dtOPDF.Rows.Count, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            '--------------------------------------------------------
            sQuery = " SELECT ""AcctNum"",""BankCode"",""CheckNum"",""CheckSum"",""Currency"",""DocNum"",""DueDate"",""LogInstanc""  FROM """ & oCompany.CompanyDB & """.""VPM1"" WHERE ""DocNum"" IN " & sListDocEntry & " "
            dtPDF1 = dsPAYMENT.Tables("PDF1")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPDF1)
            '--------------------------------------------------------
            sQuery = " SELECT ""CreditAcct"",""CreditCard"",""CreditCur"",""CreditRate"",""CreditSum"",""DocNum"",""FirstDue"",""FirstSum"",""VoucherNum"",""LogInstanc"" FROM """ & oCompany.CompanyDB & """.""VPM3"" WHERE ""DocNum"" IN " & sListDocEntry & " "
            dtPDF3 = dsPAYMENT.Tables("PDF3")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtPDF3)
            '--------------------------------------------------------
            sQuery = " SELECT ""AcctCode"",""AcctName"",""Descrip"",""DocNum"",""GrossAmnt"",""GrssAmntFC"",""VatAmnt"",""VatAmntFC"",""VatPrcnt"",""LogInstanc"",""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5""  FROM """ & oCompany.CompanyDB & """.""VPM4"" WHERE ""DocNum"" IN " & sListDocEntry & " "
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
            sQuery = "SELECT ""FaxF"",""Phone1F"",""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"", ""DdctOffice"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            dtOADM = dsPAYMENT.Tables("OADM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)

            '--------------------------------------------------------
            Select Case g_bShowDetails
                Case True
                    sQuery = "  SELECT  ""PaymentDocType"", ""PaymentDocEntry"", ""PaymentDocNum"", ""InvType"", ""InvoiceId"", ""SumApplied"", ""AppliedFC"", "
                    sQuery &= " ""PaymentDocRate"", ""PaymentObjType"", ""vatApplied"", ""vatAppldFC"", ""VisOrder"", ""DocEntry"",  "
                    sQuery &= " ""ItemCode"", ""Dscription"", ""Quantity"", ""Price"", ""LineTotal"", ""TotalFrgn"", ""ObjType"", "
                    sQuery &= " ""NumAtCard"", ""DocNum"", ""DocType"", ""DocCur"", ""DocRate"", ""DocTotal"", ""DocTotalFC"",  "
                    sQuery &= " ""VatSum"", ""VatSumFC"", ""DocDate"", ""DocDueDate"", ""TaxDate"", ""SeriesName"", ""Comments"" "
                    sQuery &= " FROM """ & oCompany.CompanyDB & """.""NCM_VIEW_RPV_INVOICE"" "
                    sQuery &= " WHERE ""PaymentDocEntry"" In " & sListDocEntry & " And ""PaymentObjType"" = '46' "
                Case False
                    sQuery = "  SELECT  ""PaymentDocType"", ""PaymentDocEntry"", ""PaymentDocNum"", ""InvType"", ""InvoiceId"", ""SumApplied"", ""AppliedFC"", "
                    sQuery &= " ""PaymentDocRate"", ""PaymentObjType"", ""vatApplied"", ""vatAppldFC"", ""VisOrder"", ""DocEntry"",  "
                    sQuery &= " ""ItemCode"", ""Dscription"", ""Quantity"", ""Price"", ""LineTotal"", ""TotalFrgn"", ""ObjType"", "
                    sQuery &= " ""NumAtCard"", ""DocNum"", ""DocType"", ""DocCur"", ""DocRate"", ""DocTotal"", ""DocTotalFC"",  "
                    sQuery &= " ""VatSum"", ""VatSumFC"", ""DocDate"", ""DocDueDate"", ""TaxDate"", ""SeriesName"", ""Comments"" "
                    sQuery &= " FROM """ & oCompany.CompanyDB & """.""NCM_VIEW_RPV_INVOICE_SUMM"" "
                    sQuery &= " WHERE ""PaymentDocEntry"" IN " & sListDocEntry & " And ""PaymentObjType"" = '46' "
            End Select

            dtVIEW = dsPAYMENT.Tables("NCM_VIEW_DRAFTPV_INVOICE")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtVIEW)

            If dtVIEW.Rows.Count > 0 Then
                SBO_Application.StatusBar.SetText("VIEW Records found : " & dtVIEW.Rows.Count, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            '--------------------------------------------------------
            dbConn.Close()

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PrepareDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Event Handlers"
    Friend Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                        Select Case pVal.ItemUID
                            Case "txtSDocNum", "txtEDocNum"
                                oEdit = oForm.Items.Item(pVal.ItemUID).Specific
                                If (oEdit.String.ToString.Trim = "") And (pVal.CharPressed = 9) Then
                                    SBO_Application.SendKeys("+{F2}")
                                    Return False
                                End If
                        End Select
                End Select
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        Select Case pVal.ItemUID
                            Case "txtSDocNum"
                                oEdit = oForm.Items.Item(pVal.ItemUID).Specific
                                Dim sDocNum As String = ""
                                Dim sDocEntry As String = "0"
                                sDocNum = oEdit.String.ToString.Trim
                                If sDocNum = "" Then
                                    oForm.DataSources.UserDataSources.Item("tbEntFr").ValueEx = 0
                                Else
                                    Dim sCheck As String = ""
                                    Dim oCheck As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    sCheck = "SELECT TOP 1 ""DocEntry"" FROM ""OVPM"" WHERE ""DocNum"" = '" & sDocNum & "' ORDER BY ""DocEntry"" DESC "
                                    oCheck.DoQuery(sCheck)
                                    If oCheck.RecordCount > 0 Then
                                        sDocEntry = oCheck.Fields.Item(0).Value
                                    End If
                                    oForm.DataSources.UserDataSources.Item("tbEntFr").ValueEx = sDocEntry
                                End If
                                oForm.Items.Item("lkEntFr").LinkTo = "tbEntFr"


                            Case "txtEDocNum"
                                oEdit = oForm.Items.Item(pVal.ItemUID).Specific
                                Dim sDocNum As String = ""
                                Dim sDocEntry As String = ""
                                sDocNum = oEdit.String.ToString.Trim
                                If sDocNum = "" Then
                                    oForm.DataSources.UserDataSources.Item("tbEntTo").ValueEx = 0
                                Else
                                    Dim sCheck As String = ""
                                    Dim oCheck As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    sCheck = "SELECT TOP 1 ""DocEntry"" FROM ""OVPM"" WHERE ""DocNum"" = '" & sDocNum & "' ORDER BY ""DocEntry"" DESC "
                                    oCheck.DoQuery(sCheck)
                                    If oCheck.RecordCount > 0 Then
                                        sDocEntry = oCheck.Fields.Item(0).Value
                                    End If
                                    oForm.DataSources.UserDataSources.Item("tbEntTo").ValueEx = sDocEntry
                                End If
                                oForm.Items.Item("lkEntTo").LinkTo = "tbEntTo"
                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Select Case pVal.ItemUID
                            Case GetItemUID(PaymentVoucherRangeItems.BTN_PRINT)
                                If oForm.Items.Item(pVal.ItemUID).Enabled Then
                                    Dim iReturn As Integer = 0

                                    If oForm.DataSources.UserDataSources.Item("ckWizard").ValueEx = "1" Then
                                        iReturn = SBO_Application.MessageBox("Please confirm if you want to list Outgoing Payments using Payment Wizard.", 2, "&Yes", "&No")
                                        If iReturn = 1 Then
                                            If IsParametersValid() Then
                                                myThread = New System.Threading.Thread(AddressOf Print_Report)
                                                myThread.SetApartmentState(Threading.ApartmentState.STA)
                                                myThread.Start()
                                            End If
                                        End If
                                    Else
                                        If IsParametersValid() Then
                                            myThread = New System.Threading.Thread(AddressOf Print_Report)
                                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                                            myThread.Start()
                                        End If
                                    End If

                                End If
                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If pVal.ItemUID = GetItemUID(PaymentVoucherRangeItems.CBO_DocType) Then
                            Dim oCombo As SAPbouiCOM.ComboBox
                            oCombo = oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.CBO_DocType)).Specific
                            Select Case oCombo.Selected.Description
                                Case "Supplier"
                                    SetupChooseFromList("S")
                                Case "Customer"
                                    SetupChooseFromList("C")
                                Case Else
                                    SetupChooseFromList("")
                            End Select
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        Dim sCFL_ID As String = ""
                        Dim sItemCode As String = ""
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim oDataTable As SAPbouiCOM.DataTable

                        oCFLEvento = pVal
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        oDataTable = oCFLEvento.SelectedObjects

                        Select Case pVal.ItemUID
                            Case GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode), GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode)
                                Try
                                    sItemCode = oDataTable.GetValue("CardCode", 0)
                                Catch ex As Exception

                                End Try
                                oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = sItemCode

                            Case "txtSBPGrp"
                                Try
                                    sItemCode = oDataTable.GetValue("GroupCode", 0)
                                Catch ex As Exception

                                End Try
                                oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = sItemCode

                            Case "txtEBPGrp"
                                Try
                                    sItemCode = oDataTable.GetValue("GroupCode", 0)
                                Catch ex As Exception

                                End Try
                                oForm.DataSources.UserDataSources.Item(pVal.ItemUID).ValueEx = sItemCode

                        End Select
                End Select
            End If
        Catch ex As Exception
            BubbleEvent = False
            SBO_Application.StatusBar.SetText("[ItemEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class

Public Enum PaymentVoucherRangeItems
    TXT_StartingDocNum = 0
    TXT_EndingDocNum = 1
    TXT_StartingBPCode = 2
    TXT_EndingBPCode = 3
    TXT_StartingDate = 4
    TXT_EndingDate = 5
    CBO_DocType = 6
    BTN_PRINT = 7
    BTN_CANCEL = 8
    CFL_StartingBPCode = 9
    CFL_EndingBPCode = 10
    CHK_IncludeCancel = 11
End Enum
Public Enum PaymentVoucherRangeDocTypes
    All = 0
    Account = 1
    Customer = 2
    Supplier = 3
End Enum