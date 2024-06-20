Imports System.IO
Imports System.Data.SqlClient

Public Class frmTMSAR

#Region "Global Variables"
    Private sQuery As String = ""
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False

    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
    Private oItem As SAPbouiCOM.Item
    Private Const FRM_UID As String = "ncm_TMSAR"
    Private Const SRF_NAME As String = "Inecom_SDK_Reporting_Package.ncmTM_StockAging.srf"

    Dim AsAtDate As DateTime = DateTime.Now
    Dim sItemFr As String = String.Empty
    Dim sItemTo As String = String.Empty
    Dim sItemGrpFr As String = String.Empty
    Dim sItemGrpTo As String = String.Empty
    Dim sWhseFr As String = String.Empty
    Dim sWhseTo As String = String.Empty
    Dim sDim1Fr As String = String.Empty
    Dim sDim1To As String = String.Empty
    Dim sDim2Fr As String = String.Empty
    Dim sDim2To As String = String.Empty
    Dim sItemTypeFr As String = String.Empty
    Dim sItemTypeTo As String = String.Empty
#End Region

#Region "Constructors"
    Public Sub New()
        MyBase.new()
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region

#Region "Initialization"
    Public Sub ShowForm()
        If LoadFromXML(SRF_NAME) Then
            oForm = SBO_Application.Forms.Item(FRM_UID)
            AddDataSource()
            oForm.Visible = True
            SetupChooseFromList()
        Else
            Try
                oForm = SBO_Application.Forms.Item(FRM_UID)
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
        Dim oCbox As SAPbouiCOM.ComboBox
        Try
            With oForm.DataSources.UserDataSources
                .Add("txtAsAtD", SAPbouiCOM.BoDataType.dt_DATE, 254)
                .Add("txtItmGrpF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtItmGrpT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtItemFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtItemTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtWhseFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtWhseTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtDim1Fr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtDim1To", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtDim2Fr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtDim2To", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtItmTypF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtItmTypT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("cbType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            End With

            oCbox = oForm.Items.Item("cbType").Specific
            oCbox.DataBind.SetBound(True, String.Empty, "cbType")
            oCbox.ValidValues.Add("0", "Original")
            oCbox.ValidValues.Add("1", "Sort By Warehouse")
            oCbox.ValidValues.Add("2", "Sort By Item")
            oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.DataSources.UserDataSources.Item("cbType").ValueEx = "0"

            oEdit = DirectCast(oForm.Items.Item("txtAsAtD").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtAsAtD")
            oForm.DataSources.UserDataSources.Item("txtAsAtD").ValueEx = DateTime.Now.ToString("yyyyMMdd")

            oEdit = oForm.Items.Item("txtItmGrpF").Specific
            oEdit.DataBind.SetBound(True, "", "txtItmGrpF")
            oEdit = oForm.Items.Item("txtItmGrpT").Specific
            oEdit.DataBind.SetBound(True, "", "txtItmGrpT")

            oEdit = oForm.Items.Item("txtItemFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtItemFr")
            oEdit = oForm.Items.Item("txtItemTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtItemTo")

            oEdit = oForm.Items.Item("txtDim1Fr").Specific
            oEdit.DataBind.SetBound(True, "", "txtDim1Fr")
            oEdit = oForm.Items.Item("txtDim1To").Specific
            oEdit.DataBind.SetBound(True, "", "txtDim1To")

            oEdit = oForm.Items.Item("txtDim2Fr").Specific
            oEdit.DataBind.SetBound(True, "", "txtDim2Fr")
            oEdit = oForm.Items.Item("txtDim2To").Specific
            oEdit.DataBind.SetBound(True, "", "txtDim2To")

            oEdit = oForm.Items.Item("txtWhseFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtWhseFr")
            oEdit = oForm.Items.Item("txtWhseTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtWhseTo")

            oEdit = oForm.Items.Item("txtItmTypF").Specific
            oEdit.DataBind.SetBound(True, "", "txtItmTypF")
            oEdit = oForm.Items.Item("txtItmTypT").Specific
            oEdit.DataBind.SetBound(True, "", "txtItmTypT")
        Catch ex As Exception
            SBO_Application.MessageBox("[frmTMSAR].[AddDataSource]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function setChooseFromListConditions(ByVal myVal As String, ByVal compareVal As String, ByVal cflUID As String) As Boolean
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFL = oForm.ChooseFromLists.Item(cflUID)
        oCons = New SAPbouiCOM.Conditions
        Try


            Select Case cflUID
                Case "cflItemFr"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    End If

                Case "cflItemTo"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "ItemCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                    End If
                Case "cflWhseFr"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                    End If

                Case "cflWhseTo"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "WhsCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                    End If

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

                Case Else
                    Throw New Exception("Invalid Choose from list. UID#" & cflUID)
                    Exit Select
            End Select
            oCFL.SetConditions(oCons)
            Return True
        Catch ex As Exception
            Throw New Exception("[frmTMSAR].[SetChooseFromList] " & vbNewLine & ex.Message)
            Return False
        End Try
        Return False
    End Function
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions = Nothing
        Dim oCon As SAPbouiCOM.Condition = Nothing
        Try
            oCFLs = oForm.ChooseFromLists

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "4"
            oCFLCreation.UniqueID = "cflItemFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtItemFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflItemFr"
            oEditLn.ChooseFromListAlias = "ItemCode"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "4"
            oCFLCreation.UniqueID = "cflItemTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtItemTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflItemTo"
            oEditLn.ChooseFromListAlias = "ItemCode"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "64"
            oCFLCreation.UniqueID = "cflWhseFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtWhseFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflWhseFr"
            oEditLn.ChooseFromListAlias = "WhsCode"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "64"
            oCFLCreation.UniqueID = "cflWhseTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtWhseTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflWhseTo"
            oEditLn.ChooseFromListAlias = "WhsCode"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "52"
            oCFLCreation.UniqueID = "cflItmGrpF"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtItmGrpF").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflItmGrpF"
            oEditLn.ChooseFromListAlias = "ItmsGrpNam"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "52"
            oCFLCreation.UniqueID = "cflItmGrpT"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtItmGrpT").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflItmGrpT"
            oEditLn.ChooseFromListAlias = "ItmsGrpNam"
        Catch ex As Exception
            Throw New Exception("[frmTMSAR].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = ""
            Select Case oForm.DataSources.UserDataSources.Item("cbType").ValueEx
                Case "0"
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_TM_V1)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
                Case "1"
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_TM_V2)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
                Case "2"
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_TM_V3)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
            End Select
            Return False
        Catch ex As Exception
            g_sReportFilename = ""
            SBO_Application.StatusBar.SetText("[TM Stock Aging].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        Try
            oForm.Items.Item("btnPrint").Enabled = False
        Catch ex As Exception

        End Try
        Try
            Dim frm As New frmTMSAR_2
            Dim bIsContinue As Boolean = False
            Try
                If ExecuteProcedure() Then
                    frm.ReportType = oForm.DataSources.UserDataSources.Item("cbType").ValueEx
                    frm.IsReportShared = g_bIsShared
                    frm.CrystalReportPath = g_sReportFilename
                    frm.AsAtDate = AsAtDate
                    frm.StartingItemGroup = sItemGrpFr
                    frm.EndingItemGroup = sItemGrpTo
                    frm.StartingItemCode = sItemFr
                    frm.EndingItemCode = sItemTo
                    frm.StartingWarehouse = sWhseFr
                    frm.EndingWarehouse = sWhseTo
                    frm.StartingDim1 = sDim1Fr
                    frm.EndingDim1 = sDim1To
                    frm.StartingDim2 = sDim2Fr
                    frm.EndingDim2 = sDim2To
                    frm.LocalCurrency = GetLocalCurrency()
                    frm.StartingItemType = sItemTypeFr
                    frm.EndingItemType = sItemTypeTo
                    bIsContinue = True
                End If
            Catch ex As Exception
                Throw ex
            Finally
                Try
                    oForm.Items.Item("btnPrint").Enabled = True
                Catch ex As Exception

                End Try
            End Try
            If bIsContinue Then
                frm.ShowDialog()
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[frmTMSAR] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function ExecuteProcedure() As Boolean
        Try
            g_bIsShared = IsSharedFileExist()

            'Get Parameter Value
            With oForm.DataSources.UserDataSources
                If .Item("txtAsAtD").ValueEx.Length > 0 Then
                    If Not DateTime.TryParseExact(.Item("txtAsAtD").ValueEx, "yyyyMMdd", Nothing, Globalization.DateTimeStyles.None, AsAtDate) Then
                        AsAtDate = DateTime.Now
                    End If
                Else
                    AsAtDate = DateTime.Now
                End If

                sItemFr = .Item("txtItemFr").ValueEx
                sItemTo = .Item("txtItemTo").ValueEx
                sWhseFr = .Item("txtWhseFr").ValueEx
                sWhseTo = .Item("txtWhseTo").ValueEx
                sDim1Fr = .Item("txtDim1Fr").ValueEx
                sDim1To = .Item("txtDim1To").ValueEx
                sDim2Fr = .Item("txtDim2Fr").ValueEx
                sDim2To = .Item("txtDim2To").ValueEx
                sItemGrpFr = .Item("txtItmGrpF").ValueEx
                sItemGrpTo = .Item("txtItmGrpT").ValueEx
                sItemTypeFr = .Item("txtItmTypF").ValueEx
                sItemTypeTo = .Item("txtItmTypT").ValueEx
            End With

            Dim sInput As String() = New String(15) {}
            Dim sr As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("Inecom_SDK_Reporting_Package.TM_SAR.sql")
            Dim bytes(sr.Length) As Byte
            sr.Position = 0
            sr.Read(bytes, 0, sr.Length)
            Dim sQueryFormat As String = System.Text.Encoding.ASCII.GetString(bytes)

            sInput(0) = sItemFr
            sInput(1) = sItemTo
            sInput(2) = sWhseFr
            sInput(3) = sWhseTo
            sInput(4) = sDim1Fr
            sInput(5) = sDim1To
            sInput(6) = sDim2Fr
            sInput(7) = sDim2To
            sInput(8) = sItemGrpFr
            sInput(9) = sItemGrpTo
            sInput(10) = AsAtDate.ToString("yyyyMMdd")
            sInput(11) = oCompany.UserName
            sInput(12) = sItemTypeFr
            sInput(13) = sItemTypeTo

            Dim sQuery As String = String.Format(sQueryFormat, sInput)
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[frmTMSAR][ExecuteProcedure]:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function ValidateParameter() As Boolean
        Try
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty

            sStart = oForm.DataSources.UserDataSources.Item("txtAsAtD").ValueEx
            If sStart.Length = 0 Then
                SBO_Application.StatusBar.SetText("Please enter as at date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oForm.ActiveItem = "txtAsAtD"
                Return False
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtItemFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtItemTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Item Code From is greater than Item Code To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtItemFr"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtItmTypF").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtItmTypT").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Item Type From is greater than Item Type To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtItmTypF"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtWhseFr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtWhseTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Warehouse From is greater than Warehouse To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtWhseFr"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtItmGrpF").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtItmGrpT").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Item Group Name From is greater than Item Group Name To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtItmGrpF"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtDim1Fr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtDim1To").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Cost Centre From is greater than Cost Centre To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtDim1Fr"
                    Return False
                End If
            End If

            sStart = oForm.DataSources.UserDataSources.Item("txtDim2Fr").ValueEx
            sEnd = oForm.DataSources.UserDataSources.Item("txtDim2To").ValueEx

            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Department From is greater than Department To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtDim2Fr"
                    Return False
                End If
            End If

            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[frmTMSAR].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function GetLocalCurrency() As String
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT MainCurncy FROM OADM")
            If (oRec.RecordCount > 0) Then
                oRec.MoveFirst()
                Return oRec.Fields.Item(0).Value.ToString()
            End If
            Return ""
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[TMSAR].[GetLocalCurrency] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End Try
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    If pVal.ItemUID = "txtDim1Fr" OrElse pVal.ItemUID = "txtDim1To" OrElse pVal.ItemUID = "txtDim2Fr" OrElse pVal.ItemUID = "txtDim2To" OrElse pVal.ItemUID = "txtItmTypF" OrElse pVal.ItemUID = "txtItmTypT" Then
                        oEdit = oForm.Items.Item(pVal.ItemUID).Specific
                        If (oEdit.Value = String.Empty) And (pVal.CharPressed = 9) Then
                            SBO_Application.SendKeys("+{F2}")
                            BubbleEvent = False
                        End If
                    End If
                End If
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    Dim sCFL_ID As String = String.Empty

                    oCFLEvento = pVal
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim myVal As String = String.Empty
                    Dim compareVal As String = String.Empty
                    Select Case pVal.ItemUID
                        Case "txtItemFr"
                            myVal = DirectCast(oForm.Items.Item("txtItemFr").Specific, SAPbouiCOM.EditText).Value
                            compareVal = DirectCast(oForm.Items.Item("txtItemTo").Specific, SAPbouiCOM.EditText).Value
                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                        Case "txtItemTo"
                            myVal = DirectCast(oForm.Items.Item("txtItemTo").Specific, SAPbouiCOM.EditText).Value
                            compareVal = DirectCast(oForm.Items.Item("txtItemFr").Specific, SAPbouiCOM.EditText).Value
                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                        Case "txtWhseFr"
                            myVal = DirectCast(oForm.Items.Item("txtWhseFr").Specific, SAPbouiCOM.EditText).Value
                            compareVal = DirectCast(oForm.Items.Item("txtWhseTo").Specific, SAPbouiCOM.EditText).Value
                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                        Case "txtWhseTo"
                            myVal = DirectCast(oForm.Items.Item("txtWhseTo").Specific, SAPbouiCOM.EditText).Value
                            compareVal = DirectCast(oForm.Items.Item("txtWhseFr").Specific, SAPbouiCOM.EditText).Value
                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                        Case "txtItmGrpF"
                            myVal = DirectCast(oForm.Items.Item("txtItmGrpF").Specific, SAPbouiCOM.EditText).Value
                            compareVal = DirectCast(oForm.Items.Item("txtItmGrpT").Specific, SAPbouiCOM.EditText).Value
                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                        Case "txtItmGrpT"
                            myVal = DirectCast(oForm.Items.Item("txtItmGrpT").Specific, SAPbouiCOM.EditText).Value
                            compareVal = DirectCast(oForm.Items.Item("txtItmGrpF").Specific, SAPbouiCOM.EditText).Value
                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
                    End Select
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pVal.ItemUID = "btnPrint" Then
                        If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                            Return ValidateParameter()
                        End If
                    End If
                End If
            Else
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then
                    Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvent = pVal
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oDataTable = oCFLEvent.SelectedObjects
                    If (Not oDataTable Is Nothing) Then
                        Dim sTemp As String = String.Empty
                        Select Case oCFLEvent.ChooseFromListUID
                            Case "cflItemFr"
                                sTemp = oDataTable.GetValue("ItemCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtItemFr").ValueEx = sTemp
                            Case "cflItemTo"
                                sTemp = oDataTable.GetValue("ItemCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtItemTo").ValueEx = sTemp
                            Case "cflWhseFr"
                                sTemp = oDataTable.GetValue("WhsCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtWhseFr").ValueEx = sTemp
                            Case "cflWhseTo"
                                sTemp = oDataTable.GetValue("WhsCode", 0)
                                oForm.DataSources.UserDataSources.Item("txtWhseTo").ValueEx = sTemp
                            Case "cflItmGrpF"
                                sTemp = oDataTable.GetValue("ItmsGrpNam", 0)
                                oForm.DataSources.UserDataSources.Item("txtItmGrpF").ValueEx = sTemp
                            Case "cflItmGrpT"
                                sTemp = oDataTable.GetValue("ItmsGrpNam", 0)
                                oForm.DataSources.UserDataSources.Item("txtItmGrpT").ValueEx = sTemp
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
            SBO_Application.StatusBar.SetText("[frmTMSAR].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class