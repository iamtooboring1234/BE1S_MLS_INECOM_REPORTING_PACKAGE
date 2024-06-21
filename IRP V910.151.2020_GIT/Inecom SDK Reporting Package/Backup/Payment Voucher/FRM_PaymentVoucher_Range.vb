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
    Private dsPAYMENT As DataSet

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

                oForm.Items.Item("20").Visible = False
                oForm.Items.Item("cbLayout").Visible = False

                InitializeItem()
                'SetupChooseFromList()
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
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams

        Try
            oCFLs = oForm.ChooseFromLists
         
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_VendorPayment
            oCFLCreation.UniqueID = "CFL_PAYMENT_FR"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtSDocNum").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_PAYMENT_FR"
            oEditLn.ChooseFromListAlias = "DocNum"
            ' ----------------------------------------
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = SAPbouiCOM.BoLinkedObject.lf_VendorPayment
            oCFLCreation.UniqueID = "CFL_PAYMENT_TO"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtEDocNum").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "CFL_PAYMENT_TO"
            oEditLn.ChooseFromListAlias = "DocNum"
            ' ----------------------------------------
        Catch ex As Exception
            Throw New Exception("[RPV].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub

    Friend Sub InitializeItem()
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oLink As SAPbouiCOM.LinkedButton

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

        With oForm.DataSources.UserDataSources
            .Add("txtSDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            .Add("txtEDocNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_StartingBPCode), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_EndingBPCode), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_StartingDate), SAPbouiCOM.BoDataType.dt_DATE, 254)
            .Add(GetItemUID(PaymentVoucherRangeItems.TXT_EndingDate), SAPbouiCOM.BoDataType.dt_DATE, 254)
            .Add(GetItemUID(PaymentVoucherRangeItems.CBO_DocType), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            .Add(GetItemUID(PaymentVoucherRangeItems.CHK_IncludeCancel), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

            .Add("tbEntFr", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
            .Add("tbEntTo", SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
        End With

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

        oEdit = oForm.Items.Item("txtSDocNum").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtSDocNum")
        oEdit = oForm.Items.Item("txtEDocNum").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtEDocNum")

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
    Private Function setChooseFromListConditions(ByVal myVal As String, ByVal compareVal As String, ByVal cflUID As String) As Boolean
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFL = oForm.ChooseFromLists.Item(cflUID)
        oCons = New SAPbouiCOM.Conditions
        Try
            Select Case cflUID
                Case GetItemUID(PaymentVoucherRangeItems.CFL_StartingBPCode)
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    End If
                Case GetItemUID(PaymentVoucherRangeItems.CFL_EndingBPCode)
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal

                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal

                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal

                    End If
                Case Else
                    Throw New Exception("Invalid Choose from list. UID#" & cflUID)
                    Exit Select
            End Select
            oCFL.SetConditions(oCons)
            Return True
        Catch ex As Exception
            Throw New Exception("[PaymentVoucherRange].[SetChooseFromList] " & vbNewLine & ex.Message)
            Return False
        End Try
        Return False
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
    Private Function IsSharedFileExist() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            g_sReportFilename = ""
            g_StructureFilename = ""

            sQuery = " SELECT IFNULL(""STRUCTUREPATH"",'') FROM ""@NCM_RPT_STRUCTURE"" "
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(ReportName.PV_Range) & "'"

            g_sReportFilename = GetSharedFilePath(ReportName.PV_Range)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    'Return True
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
            SBO_Application.StatusBar.SetText("[IRA].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub Print_Report()
        oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.BTN_PRINT)).Enabled = False
        Dim sFinalExportPath As String = ""
        Dim sFinalFileName As String = ""

        Try
            Dim frm As Hydac_FormViewer = New Hydac_FormViewer
            Dim bIsContinue As Boolean = False
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\ROP_{1}.pdf"
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
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\ROP\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\ROP_" & sCurrDate & "_" & sCurrTime & ".pdf"
            ' ===============================================================================

            Try

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

                sLoop = " SELECT ""DocNum"", ""DocEntry"" FROM """ & oCompany.CompanyDB & """.""OVPM"" "
                sLoop &= " WHERE 1=1 "

                If iIsIncludeCancel = 1 Then
                    ' sLoop &= " AND ""Canceled"" = 'Y' " ' doesnt matter Y or N
                Else
                    sLoop &= " AND ""Canceled"" = 'N'  "
                End If

                Select Case sTemp
                    Case "1"
                        sLoop &= " AND ""DocType"" = 'A' "
                    Case "2"
                        sLoop &= " AND ""DocType"" = 'C' "
                    Case "3"
                        sLoop &= " AND ""DocType"" = 'S' "
                End Select

                If sBPCodeS.Trim.Length > 0 Then
                    sLoop &= " AND ""CardCode"" >= '" & sBPCodeS.Trim & "' "
                End If
                If sBPCodeE.Trim.Length > 0 Then
                    sLoop &= " AND ""CardCode"" <= '" & sBPCodeE.Trim & "' "
                End If
                If sDocNumS.Trim.Length > 0 Then
                    sLoop &= " AND ""DocNum"" >= '" & sDocNumS.Trim & "' "
                End If
                If sDocNumE.Trim.Length > 0 Then
                    sLoop &= " AND ""DocNum"" <= '" & sDocNumE.Trim & "' "
                End If
                If sDocDateS.Trim.Length > 0 Then
                    sLoop &= " AND ""DocDate"" >= '" & sDocDateS.Trim & "' "
                End If
                If sDocDateE.Trim.Length > 0 Then
                    sLoop &= " AND ""DocDate"" <= '" & sDocDateE.Trim & "' "
                End If

                sLoop &= " GROUP BY ""DocNum"", ""DocEntry"" "

                oLoop.DoQuery(sLoop)
                If oLoop.RecordCount > 0 Then
                    Dim sListDocNum As String = "("
                    Dim sListDocEntry As String = "("
                    oLoop.MoveFirst()

                    While Not oLoop.EoF
                        sListDocNum &= "'" & oLoop.Fields.Item(0).Value & "',"
                        sListDocEntry &= "'" & oLoop.Fields.Item(1).Value & "',"

                        oLoop.MoveNext()
                    End While

                    sListDocNum = sListDocNum.Remove(sListDocNum.Length - 1, 1)
                    sListDocNum = sListDocNum & ")"

                    sListDocEntry = sListDocEntry.Remove(sListDocEntry.Length - 1, 1)
                    sListDocEntry = sListDocEntry & ")"
                    '=========================================
                    myPaymentVoucherDocType = iTemp

                    If PrepareDataset(sListDocNum, sListDocEntry) Then
                        With frm
                            Select Case SBO_Application.ClientType
                                Case SAPbouiCOM.BoClientType.ct_Desktop
                                    .ClientType = "D"
                                Case SAPbouiCOM.BoClientType.ct_Browser
                                    .ClientType = "S"
                            End Select

                            .ExportPath = sFinalFileName
                            .Dataset = dsPAYMENT
                            .Text = "Payment Voucher Range Report"
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
                        End With
                        bIsContinue = True
                    End If
                Else
                    SBO_Application.StatusBar.SetText("No data found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If

            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item(GetItemUID(PaymentVoucherRangeItems.BTN_PRINT)).Enabled = True
            End Try
            If bIsContinue Then
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
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[Print_Report] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
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
            sQuery = "SELECT  ""Block"", ""City"", ""County"",""Country"",""Code"",""State"",""ZipCode"",""Street"",""IntrntAdrs"",""LogInstanc"" FROM """ & oCompany.CompanyDB & """.""ADM1""  "
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
            sQuery = " SELECT * FROM """ & oCompany.CompanyDB & """.""OVPM"" WHERE ""DocEntry"" IN " & sListDocEntry & " "
            dtOPDF = dsPAYMENT.Tables("OPDF")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOPDF)
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
            sQuery = "SELECT ""Code"",""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""MainCurncy"",""RevOffice"",""Phone1"",""Phone2"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            dtOADM = dsPAYMENT.Tables("OADM")
            HANAcmd = dbConn.CreateCommand()
            HANAcmd.CommandText = sQuery
            HANAcmd.ExecuteNonQuery()
            HANAda.SelectCommand = HANAcmd
            HANAda.Fill(dtOADM)

            '--------------------------------------------------------
            Select Case g_bShowDetails
                Case True
                    sQuery = "SELECT * FROM """ & oCompany.CompanyDB & """.""NCM_VIEW_RPV_INVOICE"" WHERE ""PaymentDocEntry"" IN " & sListDocEntry & " AND ""PaymentObjType"" = '46' "
                Case False
                    sQuery = "SELECT * FROM """ & oCompany.CompanyDB & """.""NCM_VIEW_RPV_INVOICE_SUMM"" WHERE ""PaymentDocEntry"" IN " & sListDocEntry & " AND ""PaymentObjType"" = '46' "
            End Select

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
                                    If IsParametersValid() Then
                                        myThread = New System.Threading.Thread(AddressOf Print_Report)
                                        myThread.SetApartmentState(Threading.ApartmentState.STA)
                                        myThread.Start()
                                    End If
                                
                                End If
                        End Select

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