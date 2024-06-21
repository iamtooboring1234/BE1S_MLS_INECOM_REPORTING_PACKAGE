Option Strict Off
Option Explicit On 

Imports SAPbobsCOM
Imports System.Globalization
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Data.Common

Public Class StockAging_MOV

#Region "Global Variables"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oFormFIFO1 As SAPbouiCOM.Form

    Private sErrMsg As String
    Private lErrCode As Integer
    Private ds As DataSet

    Private dtFIFO As DataTable
    Private dtOADM As DataTable
    Private dtOITM As DataTable

    Private dtExportD As DataTable
    Private dtExportS As DataTable
    Dim bIsExportToExcel As Boolean = False
    Dim sRptType As String = String.Empty

    Private g_sReportFilename As String
    Private g_sXSDFilename As String = String.Empty
    Private g_iSecond As Integer
    Private g_bIsShared As Boolean = False
    Dim myThread As System.Threading.Thread
    Dim oStaticLn As SAPbouiCOM.StaticText
    Dim sTxtFormat As String = "txtB{0}txt"
    Dim sValFormat As String = "txtB{0}Val"
    Dim sFTxtFormat As String = "U_Bucket{0}Txt"
    Dim sFValFormat As String = "U_Bucket{0}Val"
    Dim iCount As Integer = 1
    Dim sTxtBTxt As String() = New String(10) {}
    Dim sTxtBVal As Integer() = New Integer(10) {}
    Dim sExcelPath As String = String.Empty
    Dim bIsSaveRunning As Boolean = True
    Dim bIsCancel As Boolean = False
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
            Dim oEdit As SAPbouiCOM.EditText
            Dim oCbox As SAPbouiCOM.ComboBox
            Dim oChck As SAPbouiCOM.CheckBox
            Dim oLink As SAPbouiCOM.LinkedButton
            Dim sCode1 As String = String.Empty
            Dim sCode2 As String = String.Empty
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            If LoadFromXML("Inecom_SDK_Reporting_Package." & FRM_NCM_MOV1 & ".srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading Form...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_MOV1)
                oFormFIFO1.EnableMenu(MenuID.Add, False)
                oFormFIFO1.EnableMenu(MenuID.Find, False)
                oFormFIFO1.EnableMenu(MenuID.Remove_Record, True)
                oFormFIFO1.EnableMenu(MenuID.Find, True)
                oFormFIFO1.EnableMenu(MenuID.Add, True)
                oFormFIFO1.EnableMenu(MenuID.Paste, True)
                oFormFIFO1.EnableMenu(MenuID.Copy, True)
                oFormFIFO1.EnableMenu(MenuID.Cut, True)

                AddChooseFromList()
                oStaticLn = DirectCast(oFormFIFO1.Items.Item("lbTimer").Specific, SAPbouiCOM.StaticText)

                With oFormFIFO1.DataSources.UserDataSources
                    .Add("uItemFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uItemTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uWareFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uWareTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uAsDate", SAPbouiCOM.BoDataType.dt_DATE)
                    .Add("uExcl", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cboRptType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cboGrpBy", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("tbItmGrpFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("tbItmGrpTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("tbItmGFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("tbItmGTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("chkExcel", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                    .Add("cboAgeType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

                    For iCount = 1 To 9
                        .Add(String.Format(sTxtFormat, iCount), SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                        .Add(String.Format(sValFormat, iCount), SAPbouiCOM.BoDataType.dt_LONG_NUMBER, 10)
                    Next

                End With

                '' Bind Form Controls
                Dim sTemp As String = String.Empty
                For iCount = 1 To 9
                    sTemp = String.Format(sTxtFormat, iCount)
                    oEdit = oFormFIFO1.Items.Item(sTemp).Specific
                    oEdit.DataBind.SetBound(True, String.Empty, sTemp)

                    sTemp = String.Format(sValFormat, iCount)
                    oEdit = oFormFIFO1.Items.Item(sTemp).Specific
                    oEdit.DataBind.SetBound(True, String.Empty, sTemp)
                Next
                ''-----------------------------------------------------
                oLink = oFormFIFO1.Items.Item("lkItemFr").Specific
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items
                oFormFIFO1.Items.Item("lkItemFr").LinkTo = "tbItemFr"

                oLink = oFormFIFO1.Items.Item("lkItemTo").Specific
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items
                oFormFIFO1.Items.Item("lkItemTo").LinkTo = "tbItemTo"

                oLink = oFormFIFO1.Items.Item("lkWareFr").Specific
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses
                oFormFIFO1.Items.Item("lkWareFr").LinkTo = "tbWareFr"

                oLink = oFormFIFO1.Items.Item("lkWareTo").Specific
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses
                oFormFIFO1.Items.Item("lkWareTo").LinkTo = "tbWareTo"

                oLink = oFormFIFO1.Items.Item("lkItmGrpFr").Specific
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_ItemGroups
                oFormFIFO1.Items.Item("lkItmGrpFr").LinkTo = "tbItmGFr"

                oLink = oFormFIFO1.Items.Item("lkItmGrpTo").Specific
                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_ItemGroups
                oFormFIFO1.Items.Item("lkItmGrpTo").LinkTo = "tbItmGTo"

                oEdit = oFormFIFO1.Items.Item("tbItemFr").Specific
                oEdit.DataBind.SetBound(True, "", "uItemFr")
                oEdit.ChooseFromListUID = "CFL_ItemFr"
                oEdit.ChooseFromListAlias = "ItemCode"

                oEdit = oFormFIFO1.Items.Item("tbItemTo").Specific
                oEdit.DataBind.SetBound(True, "", "uItemTo")
                oEdit.ChooseFromListUID = "CFL_ItemTo"
                oEdit.ChooseFromListAlias = "ItemCode"

                oEdit = oFormFIFO1.Items.Item("tbItmGrpFr").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "tbItmGrpFr")
                oEdit.ChooseFromListUID = "cflItmGrpF"
                oEdit.ChooseFromListAlias = "ItmsGrpNam"

                oEdit = oFormFIFO1.Items.Item("tbItmGrpTo").Specific
                oEdit.DataBind.SetBound(True, "", "tbItmGrpTo")
                oEdit.ChooseFromListUID = "cflItmGrpT"
                oEdit.ChooseFromListAlias = "ItmsGrpNam"

                oEdit = oFormFIFO1.Items.Item("tbItmGFr").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "tbItmGFr")

                oEdit = oFormFIFO1.Items.Item("tbItmGTo").Specific
                oEdit.DataBind.SetBound(True, "", "tbItmGTo")

                oEdit = oFormFIFO1.Items.Item("tbWareFr").Specific
                oEdit.DataBind.SetBound(True, "", "uWareFr")
                oEdit.ChooseFromListUID = "CFL_WareFr"
                oEdit.ChooseFromListAlias = "WhsCode"

                oEdit = oFormFIFO1.Items.Item("tbWareTo").Specific
                oEdit.DataBind.SetBound(True, "", "uWareTo")
                oEdit.ChooseFromListUID = "CFL_WareTo"
                oEdit.ChooseFromListAlias = "WhsCode"

                oEdit = oFormFIFO1.Items.Item("tbAsDate").Specific
                oEdit.DataBind.SetBound(True, "", "uAsDate")

                oChck = oFormFIFO1.Items.Item("ckExcl").Specific
                oChck.DataBind.SetBound(True, "", "uExcl")
                oChck.ValOff = "N"
                oChck.ValOn = "Y"

                oChck = oFormFIFO1.Items.Item("chkExcel").Specific
                oChck.DataBind.SetBound(True, String.Empty, "chkExcel")
                oChck.ValOff = "N"
                oChck.ValOn = "Y"

                oCbox = oFormFIFO1.Items.Item("cboRptType").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cboRptType")
                oCbox.ValidValues.Add("0", "Summary")
                oCbox.ValidValues.Add("1", "Detail")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormFIFO1.Items.Item("cboGrpBy").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cboGrpBy")
                oCbox.ValidValues.Add("0", "Item Code")
                oCbox.ValidValues.Add("1", "Item Code, Warehouse")
                oCbox.ValidValues.Add("2", "Warehouse, Item Code")
                oCbox.ValidValues.Add("3", "Item Group")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                oCbox = oFormFIFO1.Items.Item("cboAgeType").Specific
                oCbox.DataBind.SetBound(True, String.Empty, "cboAgeType")
                oCbox.ValidValues.Add("0", "Receipt Date")
                oCbox.ValidValues.Add("1", "Transaction Date")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                'V910.148.2023 - Defaulted to Y - MEDQUEST
                oFormFIFO1.DataSources.UserDataSources.Item("uExcl").ValueEx = "Y"

                PopulateDate()
                SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                oFormFIFO1.Visible = True
            Else
                ' Loading .srf file failed most likely it is because the form is already opened
                Try
                    oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_MOV1)
                    If oFormFIFO1.Visible Then
                        oFormFIFO1.Select()
                    Else
                        oFormFIFO1.Close()
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadForm] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddChooseFromList()
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = oFormFIFO1.ChooseFromLists
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items

            oCFLCreationParams.UniqueID = "CFL_ItemFr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            '' -----------------------------------------------------------            
            '' // InvntItem = Y, ManSerNum = N and ManBtchNum = N
            oCon = oCons.Add
            'oCon.BracketOpenNum = 2
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            'oCon.BracketCloseNum = 1
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            'oCon = oCons.Add
            'oCon.BracketOpenNum = 1
            'oCon.Alias = "EvalSystem"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "A"
            'oCon.BracketCloseNum = 2
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            'oCon = oCons.Add
            'oCon.BracketOpenNum = 2
            'oCon.Alias = "ManBtchNum"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCon.BracketCloseNum = 3
            '' -----------------------------------------------------------            
            oCFL.SetConditions(oCons)

            oCFLCreationParams.UniqueID = "CFL_ItemTo"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            '' -----------------------------------------------------------            
            '' // InvntItem = Y, ManSerNum = N and ManBtchNum = N
            oCon = oCons.Add
            'oCon.BracketOpenNum = 2
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            ' oCon.BracketCloseNum = 1
            ' oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            'oCon = oCons.Add
            'oCon.BracketOpenNum = 1
            'oCon.Alias = "EvalSystem"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "A"
            'oCon.BracketCloseNum = 2
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            'oCon = oCons.Add
            'oCon.BracketOpenNum = 2
            'oCon.Alias = "ManBtchNum"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCon.BracketCloseNum = 3
            '' -----------------------------------------------------------            
            oCFL.SetConditions(oCons)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_Warehouses

            oCFLCreationParams.UniqueID = "CFL_WareFr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.UniqueID = "CFL_WareTo"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ItemGroups
            oCFLCreationParams.UniqueID = "cflItmGrpF"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.UniqueID = "cflItmGrpT"
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ItemGroups
            oCFL = oCFLs.Add(oCFLCreationParams)


        Catch ex As Exception
            MsgBox("[Add_ChooseFromList] : " & ex.ToString)
        End Try
    End Sub
    Private Sub PopulateDate()
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sCurrDate As String = String.Empty
            Dim sQuery As String = String.Empty

            sCurrDate = GetCurrentDate()
            oFormFIFO1.DataSources.UserDataSources.Item("uAsDate").ValueEx = sCurrDate

            sTxtBTxt(0) = "0-30 Days"
            sTxtBTxt(1) = "31-60 Days"
            sTxtBTxt(2) = "61-90 Days"
            sTxtBTxt(3) = "91-120 Days"
            sTxtBTxt(4) = "121-150 Days"
            sTxtBTxt(5) = "151-180 Days"
            sTxtBTxt(6) = "181-270 Days"
            sTxtBTxt(7) = "271-365 Days"
            sTxtBTxt(8) = ">365 Days"

            sTxtBVal(0) = 30
            sTxtBVal(1) = 60
            sTxtBVal(2) = 90
            sTxtBVal(3) = 120
            sTxtBVal(4) = 150
            sTxtBVal(5) = 180
            sTxtBVal(6) = 270
            sTxtBVal(7) = 365
            sTxtBVal(8) = 365

            Dim sTemp1 As String = String.Empty
            sQuery = " SELECT * FROM ""@NCM_BUCKET"" WHERE ""U_Type"" = 'NCM_SAR_MOV1'"
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                iCount = 0
                For iCount = 0 To 8
                    sTemp1 = String.Format(sFTxtFormat, iCount + 1)
                    sTxtBTxt(iCount) = oRecord.Fields.Item(sTemp1).Value

                    sTemp1 = String.Format(sFValFormat, iCount + 1)
                    sTxtBVal(iCount) = oRecord.Fields.Item(sTemp1).Value
                Next
            End If

            For iCount = 0 To 8
                sTemp1 = String.Format(sTxtFormat, iCount + 1)
                oFormFIFO1.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBTxt(iCount)

                sTemp1 = String.Format(sValFormat, iCount + 1)
                oFormFIFO1.DataSources.UserDataSources.Item(sTemp1).ValueEx = sTxtBVal(iCount)
            Next

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PopulateDate] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub OpenSaveFileDialog()
        Try
            Dim frmSave As New frmSaveDialog
            frmSave.TopMost = True
            frmSave.Show()
            System.Threading.Thread.CurrentThread.Sleep(10)
            frmSave.Activate()

            If (frmSave.SaveFileDialog1.ShowDialog() = DialogResult.OK) Then
                sExcelPath = frmSave.SaveFileDialog1.FileName
            Else
                bIsCancel = True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            bIsSaveRunning = False
        End Try
    End Sub
#End Region

#Region "Print Report"
    Private Function ExecuteProcedure() As Boolean
        Try
            Dim sQuery As String = ""
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            With oFormFIFO1.DataSources.UserDataSources
                SBO_Application.StatusBar.SetText("Executing Stored Procedure...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                sQuery = "CALL NCM_SP_SAR_MOV1_ECS ("
                sQuery &= "'" & .Item("uAsDate").ValueEx & "', "
                sQuery &= "'" & oCompany.UserSignature & "', "
                sQuery &= "'" & .Item("uItemFr").ValueEx.Replace("'", "''") & "', "
                sQuery &= "'" & .Item("uItemTo").ValueEx.Replace("'", "''") & "', "
                sQuery &= "'" & .Item("uWareFr").ValueEx.Replace("'", "''") & "', "
                sQuery &= "'" & .Item("uWareTo").ValueEx.Replace("'", "''") & "',  "
                sQuery &= "'" & .Item("tbItmGrpFr").ValueEx.Replace("'", "''") & "', "
                sQuery &= "'" & .Item("tbItmGrpTo").ValueEx.Replace("'", "''") & "')"
            End With
            Try
                SBO_Application.StatusBar.SetText("Executing Stored Procedure...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oRec.DoQuery(sQuery)
                oRec = Nothing
                Return True
            Catch ex As Exception
                SBO_Application.StatusBar.SetText("SP is not completed successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Throw ex
            End Try
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[ExecProcedure] : " & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Private Function GenerateRecords() As Boolean
        Try
            Dim oQuery As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oExecute As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            Dim inputString As String() = New String(12) {}
            Dim ProviderName As String = "System.Data.Odbc"
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory

            SBO_Application.StatusBar.SetText("Filtering Data...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            bIsExportToExcel = IIf(oFormFIFO1.DataSources.UserDataSources.Item("chkExcel").ValueEx = "Y", True, False)

            Dim oComboLn As SAPbouiCOM.ComboBox
            oComboLn = oFormFIFO1.Items.Item("cboRptType").Specific

            If (Not oComboLn.Selected Is Nothing) Then
                sRptType = oComboLn.Selected.Value
            Else
                sRptType = "0"
            End If

            ds = New SAR_FIFO1
            With oFormFIFO1.DataSources.UserDataSources
                inputString(0) = .Item("uAsDate").ValueEx
                inputString(1) = oCompany.UserSignature.ToString()
                inputString(2) = .Item("uItemFr").ValueEx.Replace("'", "''")
                inputString(3) = .Item("uItemTo").ValueEx.Replace("'", "''")
                inputString(4) = .Item("uWareFr").ValueEx.Replace("'", "''")
                inputString(5) = .Item("uWareTo").ValueEx.Replace("'", "''")
                inputString(6) = .Item("tbItmGrpFr").ValueEx.Replace("'", "''")
                inputString(7) = .Item("tbItmGrpTo").ValueEx.Replace("'", "''")
                inputString(8) = .Item("cboAgeType").ValueEx
            End With

            dtFIFO = ds.Tables("DS_FIFO")
            dtExportS = ds.Tables("Excel_Output")
            SBO_Application.StatusBar.SetText("Inserting Data...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            '' ----------------------------------------------------------------------

            'CALL NCM_SP_SAR_MOV1 ('2014-02-15','manager','','','','','','',0)

            sQuery = "CALL NCM_SP_SAR_MOV1 ("
            sQuery &= "'" & inputString(0) & "',"
            sQuery &= "'" & inputString(1) & "',"
            sQuery &= "'" & inputString(2) & "',"
            sQuery &= "'" & inputString(3) & "',"
            sQuery &= "'" & inputString(4) & "',"
            sQuery &= "'" & inputString(5) & "',"
            sQuery &= "'" & inputString(6) & "',"
            sQuery &= "'" & inputString(7) & "',"
            sQuery &= "'" & inputString(8) & "')"
            oQuery = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oQuery.DoQuery(sQuery)
            '' ----------------------------------------------------------------------

            SBO_Application.StatusBar.SetText("Filtering Data...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim sTemp1 As String = String.Empty
            iCount = 0
            For iCount = 1 To 9 Step 1
                sTemp1 = String.Format(sValFormat, iCount)
                inputString(iCount) = oFormFIFO1.DataSources.UserDataSources.Item(sTemp1).ValueEx
            Next

            SBO_Application.StatusBar.SetText("Filling Datasets...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            iCount = 1

            Select Case oFormFIFO1.DataSources.UserDataSources.Item("cboRptType").ValueEx
                Case "0"    'SUMMARY
                    sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_SP_SAR_MOV2_SUMM"" ("
                Case Else   'DETAILS
                    sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_SP_SAR_MOV2"" ("
            End Select

            sQuery &= "'" & inputString(0) & "',"
            sQuery &= "'" & oCompany.UserSignature & "',"
            sQuery &= "'" & inputString(1) & "',"
            sQuery &= "'" & inputString(2) & "',"
            sQuery &= "'" & inputString(3) & "',"
            sQuery &= "'" & inputString(4) & "',"
            sQuery &= "'" & inputString(5) & "',"
            sQuery &= "'" & inputString(6) & "',"
            sQuery &= "'" & inputString(7) & "',"
            sQuery &= "'" & inputString(8) & "',"
            sQuery &= "'" & inputString(9) & "')"

            oExecute = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oExecute.DoQuery(sQuery)

            If oExecute.RecordCount > 0 Then

                _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
                dbConn = _DbProviderFactoryObject.CreateConnection()
                dbConn.ConnectionString = connStr
                dbConn.Open()

                Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
                Dim HANAcmd As DbCommand = dbConn.CreateCommand()
                HANAcmd.CommandText = sQuery
                HANAcmd.ExecuteNonQuery()
                HANAda.SelectCommand = HANAcmd
                HANAda.Fill(dtFIFO)

                Try
                    ' Added since V910.148.2023...
                    sQuery = "  SELECT  T1.""ItemCode"", T1.""ItemName"", IFNULL(T1.""InvntryUom"",'') ""InvntryUom"", "
                    sQuery &= "         IFNULL(T1.""U_MQ_BU"",'') AS ""U_MQ_BU"", "
                    sQuery &= "         IFNULL(T1.""U_MQ_Supplier"",'') AS ""U_MQ_SUPPLIER"", "
                    sQuery &= "         IFNULL(T1.""U_MQ_Brand"",'') AS ""U_MQ_BRAND"", "
                    sQuery &= "         IFNULL(T2.""Name"",'') AS ""U_MQ_BRANDNAME"" "
                    sQuery &= " FROM    """ & oCompany.CompanyDB & """.""OITM"" T1 "
                    sQuery &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""@MQ_ITM_BRAND"" T2 ON T1.""U_MQ_Brand"" = T2.""Code"" "

                    'sQuery = "  SELECT  ""ItemCode"", ""ItemName"", IFNULL(""InvntryUom"",'') ""InvntryUom"", "
                    'sQuery &= "         '' ""U_MQ_BU"", IFNULL(""U_MQ_Supplier"",'') AS ""U_MQ_SUPPLIER"", "
                    'sQuery &= "         '' ""U_MQ_BRAND"", '' ""U_MQ_BRANDNAME"" "
                    'sQuery &= " FROM    """ & oCompany.CompanyDB & """.""OITM"" "

                    dtOITM = ds.Tables("OITM")
                    HANAcmd = dbConn.CreateCommand()
                    HANAcmd.CommandText = sQuery
                    HANAcmd.ExecuteNonQuery()
                    HANAda.SelectCommand = HANAcmd
                    HANAda.Fill(dtOITM)

                Catch ex As Exception

                    sQuery = "  SELECT  ""ItemCode"", ""ItemName"", IFNULL(""InvntryUom"",'') ""InvntryUom"", "
                    sQuery &= "         '' ""U_MQ_BU"", ""U_MQ_SUPPLIER"", '' ""U_MQ_BRAND"", '' ""U_MQ_BRANDNAME"" "
                    sQuery &= " FROM    """ & oCompany.CompanyDB & """.""OITM"" "

                    dtOITM = ds.Tables("OITM")
                    HANAcmd = dbConn.CreateCommand()
                    HANAcmd.CommandText = sQuery
                    HANAcmd.ExecuteNonQuery()
                    HANAda.SelectCommand = HANAcmd
                    HANAda.Fill(dtOITM)

                End Try

                '--------------------------------------------------------
                'OADM (Company Details)
                '--------------------------------------------------------
                sQuery = "SELECT ""CompnyAddr"",""CompnyName"", IFNULL(""E_Mail"",'') ""E_Mail"", IFNULL(""Fax"",'') ""Fax"", IFNULL(""FreeZoneNo"",'') ""FreeZoneNo"", IFNULL(""RevOffice"",'') ""RevOffice"", IFNULL(""Phone1"",'') ""Phone1"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
                dtOADM = ds.Tables("OADM")
                HANAcmd = dbConn.CreateCommand()
                HANAcmd.CommandText = sQuery
                HANAcmd.ExecuteNonQuery()
                HANAda.SelectCommand = HANAcmd
                HANAda.Fill(dtOADM)

                dbConn.Close()
            End If
            '' ----------------------------------------------------------------------
            oExecute = Nothing
            oQuery = Nothing
            Return True

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GenerateRecords] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            Select Case oFormFIFO1.DataSources.UserDataSources.Item("cboRptType").ValueEx
                Case "0"
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_MOVAVG_SUMMARY)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
                Case Else
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_MOVAVG_DETAILS)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
            End Select
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[StockAgeing_MOVAVG].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub PrintSAR_FIFO_NonBatch()
        Try
            Dim bIsContinue As Boolean = False
            Dim oSARVwr As New SARVwr
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""

            oFormFIFO1.Items.Item("btPrint").Enabled = False

            Try
                Dim sTempDirectory As String = ""
                Dim sPathFormat As String = "{0}\STOCK_{1}.pdf"
                Dim sCurrDate As String = ""
                Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")
                Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                oRec.DoQuery("SELECT  TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    sCurrDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
                End If

                ' ===============================================================================
                ' get the folder of the current DB Name --> set to local

                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\STOCK\" & oCompany.CompanyDB
                Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                If Not di.Exists Then
                    di.Create()
                End If
                sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                sFinalFileName = di.FullName & "\STOCK_" & sCurrDate & "_" & sCurrTime & ".pdf"
                ' ===============================================================================

                g_iSecond = 0
                g_bIsShared = IsSharedFileExist()

                If ExecuteProcedure() Then
                    If GenerateRecords() Then
                        '' To print Stock Aging Report FIFO Non-Batch Items
                        SBO_Application.StatusBar.SetText("Printing...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        With oSARVwr
                            Select Case SBO_Application.ClientType
                                Case SAPbouiCOM.BoClientType.ct_Desktop
                                    .ClientType = "D"
                                Case SAPbouiCOM.BoClientType.ct_Browser
                                    .ClientType = "S"
                            End Select

                            .ExportPath = sFinalFileName
                            .Server = oCompany.Server
                            .Database = oCompany.CompanyDB
                            .DBUsername = DBUsername
                            .DBPassword = DBPassword
                            .Dataset = ds
                            .ItemFrom = oFormFIFO1.DataSources.UserDataSources.Item("uItemFr").ValueEx
                            .ItemTo = oFormFIFO1.DataSources.UserDataSources.Item("uItemTo").ValueEx
                            .WarehouseFrom = oFormFIFO1.DataSources.UserDataSources.Item("uWareFr").ValueEx
                            .WarehouseTo = oFormFIFO1.DataSources.UserDataSources.Item("uWareTo").ValueEx
                            .ExcludeZeroBalance = oFormFIFO1.DataSources.UserDataSources.Item("uExcl").ValueEx.ToString.Trim
                            .AsAtDate = oFormFIFO1.DataSources.UserDataSources.Item("uAsDate").ValueEx
                            .IsShared = g_bIsShared
                            .AgeType = oFormFIFO1.DataSources.UserDataSources.Item("cboAgeType").ValueEx
                            .BucketText = sTxtBTxt
                            .ExcelFilePath = sExcelPath
                            .IsExcel = bIsExportToExcel
                            .SharedReportName = g_sReportFilename
                            .ItemGroupFrom = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
                            .ItemGroupTo = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx

                            Select Case oFormFIFO1.DataSources.UserDataSources.Item("cboRptType").ValueEx
                                Case "0"
                                    .ReportType = "S"
                                    .ReportName = ReportName.SAR_MOVAVG_SUMMARY
                                Case Else
                                    .ReportType = "D"
                                    .ReportName = ReportName.SAR_MOVAVG_DETAILS
                            End Select

                            Select Case oFormFIFO1.DataSources.UserDataSources.Item("cboGrpBy").ValueEx
                                Case "0"
                                    .GroupBy = "ItemCode"
                                Case "1"
                                    .GroupBy = "ItemCode_Whse"
                                Case "2"
                                    .GroupBy = "WhsCode"
                                Case "3"
                                    .GroupBy = "ItemGroup"
                                Case Else
                                    .GroupBy = "OTHERS"
                            End Select
                            bIsContinue = True
                        End With
                    End If
                End If
            Catch ex As Exception
                SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                oFormFIFO1.Items.Item("btPrint").Enabled = True
                g_iSecond = 0
                oStaticLn.Caption = String.Empty
            End Try

            If bIsContinue Then
                If (bIsExportToExcel) Then
                    'Export To Excel
                    '--------------------------------------------------------------------------------
                    If (String.Compare(sRptType, "0", True) = 0) Then
                        oSARVwr.OpenSummaryReport_MOV()
                    Else
                        oSARVwr.OpenDetailReport_MOV()
                    End If
                Else

                    Select Case SBO_Application.ClientType
                        Case SAPbouiCOM.BoClientType.ct_Desktop
                            oSARVwr.ShowDialog()

                        Case SAPbouiCOM.BoClientType.ct_Browser
                            If (String.Compare(sRptType, "0", True) = 0) Then
                                oSARVwr.OpenSummaryReport_MOV()
                            Else
                                oSARVwr.OpenDetailReport_MOV()
                            End If

                            If File.Exists(sFinalFileName) Then
                                SBO_Application.SendFileToBrowser(sFinalFileName)
                            End If
                    End Select
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Sub
    Private Function ValidateParameters() As Boolean
        oFormFIFO1.ActiveItem = "tbItemFr"
        Dim sFromValue As String = String.Empty
        Dim sToValue As String = String.Empty

        Try
            sFromValue = oFormFIFO1.DataSources.UserDataSources.Item("uItemFr").ValueEx
            sToValue = oFormFIFO1.DataSources.UserDataSources.Item("uItemTo").ValueEx
            If (Not (sFromValue.Length = 0 AndAlso sToValue.Length = 0)) Then
                If (String.Compare(sFromValue, sToValue, True) > 0) Then
                    oFormFIFO1.ActiveItem = "tbItemFr"
                    SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - item from is greater than item to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            sFromValue = oFormFIFO1.DataSources.UserDataSources.Item("uWareFr").ValueEx
            sToValue = oFormFIFO1.DataSources.UserDataSources.Item("uWareTo").ValueEx
            If (Not (sFromValue.Length = 0 AndAlso sToValue.Length = 0)) Then
                If (String.Compare(sFromValue, sToValue, True) > 0) Then
                    oFormFIFO1.ActiveItem = "tbWareFr"
                    SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - warehouse from is greater than warehouse to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            sFromValue = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
            sToValue = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx
            If (Not (sFromValue.Length = 0 AndAlso sToValue.Length = 0)) Then
                If (String.Compare(sFromValue, sToValue, True) > 0) Then
                    oFormFIFO1.ActiveItem = "tbItmGrpFr"
                    SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - item group from is greater than item group to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

            Dim sTemp1 As String = String.Empty
            Dim sTemp2 As String = String.Empty
            iCount = 0
            For iCount = 0 To 8 Step 1
                sTemp1 = String.Format(sTxtFormat, iCount + 1)
                sTemp2 = String.Format(sValFormat, iCount + 1)
                sTxtBTxt(iCount) = oFormFIFO1.DataSources.UserDataSources.Item(sTemp1).ValueEx
                sTxtBVal(iCount) = Integer.Parse(oFormFIFO1.DataSources.UserDataSources.Item(sTemp2).ValueEx)
            Next

            Dim i1 As Integer = 0
            Dim i2 As Integer = 0

            For iCount = 1 To 8 Step 1
                i1 = sTxtBVal(iCount - 1)
                i2 = sTxtBVal(iCount)
                sTemp1 = String.Format(sValFormat, iCount)
                sTemp2 = String.Format(sValFormat, iCount + 1)
                If (i1 < 0) Then
                    oFormFIFO1.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - Value in bucket " & iCount.ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i2 < 0) Then
                    oFormFIFO1.ActiveItem = sTemp2
                    SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - Value in bucket " & (iCount + 1).ToString() & " is less than zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (i1 > i2) Then
                    oFormFIFO1.ActiveItem = sTemp1
                    SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - Value in bucket " & (iCount).ToString() & " is greater than value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If

                If (iCount = 8) Then
                    If (i1 <> i2) Then
                        oFormFIFO1.ActiveItem = sTemp1
                        SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - Value in bucket " & (iCount).ToString() & " is not equal to value in bucket " & (iCount + 1).ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next

            Dim sFileDialog As New SaveFileDialog
            sFileDialog.Filter = "Tab Separated Value|*.xls"
            sFileDialog.Title = "Export To"
            sFileDialog.RestoreDirectory = True
            sFileDialog.CheckFileExists = True
            sFileDialog.DefaultExt = "xls"
            sExcelPath = String.Empty
            bIsExportToExcel = IIf(oFormFIFO1.DataSources.UserDataSources.Item("chkExcel").ValueEx = "Y", True, False)

            If (bIsExportToExcel) Then
                bIsSaveRunning = True
                bIsCancel = False
                Dim myThread2 As New System.Threading.Thread(AddressOf OpenSaveFileDialog)
                myThread2.SetApartmentState(Threading.ApartmentState.STA)
                myThread2.Start()
                myThread2.Join()
                While (bIsSaveRunning)
                    System.Threading.Thread.CurrentThread.Sleep(900)
                End While
                Return Not bIsCancel
            End If

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Event Handlers"
    Private Function setChooseFromListConditions(ByVal myVal As String, ByVal compareVal As String, ByVal cflUID As String) As Boolean
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFL = oFormFIFO1.ChooseFromLists.Item(cflUID)
        oCons = New SAPbouiCOM.Conditions
        Try
            Select Case cflUID
                Case "cflItmGrpF"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    End If

                    Exit Select
                Case "cflItmGrpT"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItmsGrpNam"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                    End If
                    Exit Select
                Case Else
                    Throw New Exception("Invalid Choose from list. UID#" & cflUID)
                    Exit Select
            End Select
            oCFL.SetConditions(oCons)
            Return True
        Catch ex As Exception
            Throw New Exception("[StockAging].[SetChooseFromList] " & vbNewLine & ex.Message)
            Return False
        End Try
        Return False
    End Function
    Friend Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "btPrint" Then
                            BubbleEvent = ValidateParameters()
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        Dim sCFL_ID As String = String.Empty

                        oCFLEvento = pVal
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        Dim myVal As String = String.Empty
                        Dim compareVal As String = String.Empty
                        Select Case pVal.ItemUID
                            Case "tbItmGrpFr"
                                myVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
                                compareVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx
                                Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                            Case "tbItmGrpTo"
                                myVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx
                                compareVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
                                Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                        End Select
                End Select
            End If

            If pVal.Before_Action = False Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Select Case pVal.ItemUID
                            Case "btPrint"
                                If oFormFIFO1.Items.Item(pVal.ItemUID).Enabled Then
                                    myThread = New System.Threading.Thread(AddressOf PrintSAR_FIFO_NonBatch)
                                    myThread.SetApartmentState(Threading.ApartmentState.STA)
                                    myThread.Start()
                                End If
                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Select Case pVal.ItemUID
                            Case "tbItemFr", "tbItemTo"
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim sCFL_ID As String = ""
                                Dim sItemCode As String = ""
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oDataTable As SAPbouiCOM.DataTable

                                oCFLEvento = pVal
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
                                oDataTable = oCFLEvento.SelectedObjects

                                Try
                                    sItemCode = oDataTable.GetValue("ItemCode", 0)
                                Catch ex As Exception

                                End Try
                                Select Case pVal.ItemUID
                                    Case "tbItemFr"
                                        oFormFIFO1.DataSources.UserDataSources.Item("uItemFr").ValueEx = sItemCode
                                    Case "tbItemTo"
                                        oFormFIFO1.DataSources.UserDataSources.Item("uItemTo").ValueEx = sItemCode
                                End Select
                            Case "tbItmGrpFr", "tbItmGrpTo"
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim sCFL_ID As String = ""
                                Dim sItemGrpCod As String = ""
                                Dim sItemGrpName As String = ""
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oDataTable As SAPbouiCOM.DataTable

                                oCFLEvento = pVal
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
                                oDataTable = oCFLEvento.SelectedObjects

                                Try
                                    sItemGrpName = oDataTable.GetValue("ItmsGrpNam", 0)
                                    sItemGrpCod = oDataTable.GetValue("ItmsGrpCod", 0)
                                Catch ex As Exception

                                End Try
                                Select Case pVal.ItemUID
                                    Case "tbItmGrpFr"
                                        oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx = sItemGrpName
                                        oFormFIFO1.DataSources.UserDataSources.Item("tbItmGFr").ValueEx = sItemGrpCod

                                    Case "tbItmGrpTo"
                                        oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx = sItemGrpName
                                        oFormFIFO1.DataSources.UserDataSources.Item("tbItmGTo").ValueEx = sItemGrpCod
                                End Select
                            Case "tbWareFr", "tbWareTo"
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim sCFL_ID As String = ""
                                Dim sWareCode As String = ""
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oDataTable As SAPbouiCOM.DataTable

                                oCFLEvento = pVal
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
                                oDataTable = oCFLEvento.SelectedObjects

                                Try
                                    sWareCode = oDataTable.GetValue("WhsCode", 0)
                                Catch ex As Exception

                                End Try
                                Select Case pVal.ItemUID
                                    Case "tbWareFr"
                                        oFormFIFO1.DataSources.UserDataSources.Item("uWareFr").ValueEx = sWareCode
                                    Case "tbWareTo"
                                        oFormFIFO1.DataSources.UserDataSources.Item("uWareTo").ValueEx = sWareCode
                                End Select
                        End Select
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                        Select Case pVal.ItemUID
                            Case "txtB8Val"
                                If (pVal.ItemChanged) Then
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtB9Val").ValueEx = oFormFIFO1.DataSources.UserDataSources.Item("txtB8Val").ValueEx
                                End If
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
