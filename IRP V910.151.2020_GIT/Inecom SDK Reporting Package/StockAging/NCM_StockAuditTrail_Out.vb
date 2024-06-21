Option Strict Off
Option Explicit On 

Imports SAPbobsCOM

Public Class NCM_StockAuditTrail_Out

#Region "Global Variables"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oFormFIFO1 As SAPbouiCOM.Form
    Private sErrMsg As String
    Private lErrCode As Integer
    Private ds As DataSet

    Private oLnObj As SAPbouiCOM.LinkedButton
    Private oLnItm As SAPbouiCOM.Item
    Private oGridCl As SAPbouiCOM.Grid
    Private g_iSecond As Integer
    Private g_bIsShared As Boolean = False
    Private g_bIsXSDShared As Boolean = False
    Private oStaticLn As SAPbouiCOM.StaticText
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
            Dim oBttn As SAPbouiCOM.Button

            If LoadFromXML("Inecom_SDK_Reporting_Package." & FRM_NCM_SES2 & ".srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading Form...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_SES2)
                oGridCl = oFormFIFO1.Items.Item("mxGrid").Specific
                oBttn = oFormFIFO1.Items.Item("btPrint").Specific
                oBttn.Caption = "OK"

                oFormFIFO1.DataSources.DataTables.Add("abc")
                oGridCl.DataTable = oFormFIFO1.DataSources.DataTables.Item("abc")

                oFormFIFO1.DataSources.UserDataSources.Add("txtLink", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                oEdit = oFormFIFO1.Items.Item("txtLink").Specific
                oEdit.DataBind.SetBound(True, String.Empty, "txtLink")

                oLnItm = oFormFIFO1.Items.Item("lbObj")
                oLnObj = oFormFIFO1.Items.Item("lbObj").Specific

                PopulateData()
                oFormFIFO1.Settings.EnableRowFormat = False
                oFormFIFO1.Settings.Enabled = False
                SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                oFormFIFO1.Visible = True
            Else
                Try
                    oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_SES2)
                    PopulateData()
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
    Private Sub PopulateData()
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            sQuery = "CALL NCM_SP_SES_MOV2 ('" & oCompany.UserSignature & "')"
            oGridCl.DataTable = oFormFIFO1.DataSources.DataTables.Item("abc")
            oFormFIFO1.DataSources.DataTables.Item(0).ExecuteQuery(sQuery)
            SetGrid(oGridCl)
            oGridCl.CollapseLevel = 3
            oGridCl.AutoResizeColumns()
            oFormFIFO1.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PopulateDate] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub SetGrid(ByVal oGridLn As SAPbouiCOM.Grid)
        Dim i As Integer = 0
        Dim sColumnUID As String = ""
        Dim bVisible As Boolean = False
        Dim oEditColumn As SAPbouiCOM.EditTextColumn

        For i = 0 To oGridLn.Columns.Count - 1
            oEditColumn = oGridLn.Columns.Item(i)
            If (oEditColumn.UniqueID.Length > 2) Then
                If (String.Compare("|", oEditColumn.UniqueID.Substring(oEditColumn.UniqueID.Length - 2, 1), True) = 0) Then
                    sColumnUID = oEditColumn.UniqueID.Substring(0, oEditColumn.UniqueID.Length - 2)
                    bVisible = IIf((String.Compare("V", oEditColumn.UniqueID.Substring(oEditColumn.UniqueID.Length - 1, 1), True) = 0), True, False)
                Else
                    sColumnUID = oEditColumn.UniqueID
                    bVisible = False
                End If
            Else
                sColumnUID = oEditColumn.UniqueID
                bVisible = False
            End If

            'sColumnUID = oEditColumn.UniqueID
            'bVisible = True

            oEditColumn.Editable = False
            Select Case sColumnUID
                Case "ItmCod" ' Or "ITMCOD"
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oItems
                    oEditColumn.TitleObject.Caption = "Item No."
                    oEditColumn.Visible = bVisible
                Case "WhsCod" ' Or "WHSCOD"
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oWarehouses
                    oEditColumn.TitleObject.Caption = "Whse"
                    oEditColumn.Visible = bVisible
                Case "ItmGrp" 'Or "ITMGRP"
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oItemGroups
                    oEditColumn.TitleObject.Caption = "Item Grp."
                    oEditColumn.Visible = bVisible
                Case "TransQty" ' Or "TRANSQTY"
                    oEditColumn.TitleObject.Caption = "Quantity"
                    oEditColumn.Visible = bVisible
                Case "TransValue" 'Or "TRANSVALUE"
                    oEditColumn.TitleObject.Caption = "Trans. Value"
                    oEditColumn.Visible = bVisible
                Case "TransDate" 'Or "TRANSDATE"
                    oEditColumn.TitleObject.Caption = "Posting Date"
                    oEditColumn.Visible = bVisible
                Case "TransNam" ' Or "TRANSNAM" ''LINK
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oItemGroups
                    oEditColumn.TitleObject.Caption = "Document"
                    oEditColumn.Visible = bVisible
                Case "BaseDate" ' Or "BASEDATE"
                    oEditColumn.TitleObject.Caption = "Base. Posting Date"
                    oEditColumn.Visible = bVisible
                Case "BaseNam" 'Or "BASENAM" ''LINK
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oItemGroups
                    oEditColumn.TitleObject.Caption = "Transaction"
                    oEditColumn.Visible = bVisible
                Case "OrigDate" 'Or "ORIGDATE"
                    oEditColumn.TitleObject.Caption = "Orig. Posting Date"
                    oEditColumn.Visible = bVisible
                Case "OrigNam" 'Or "ORIGNAM" ''LINK
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oItemGroups
                    oEditColumn.TitleObject.Caption = "Orig. Transaction"
                    oEditColumn.Visible = bVisible
                Case "RefOutNam" 'Or "REFOUTNAM" ''LINK
                    oEditColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oItemGroups
                    oEditColumn.TitleObject.Caption = "Reference"
                    oEditColumn.Visible = bVisible
                Case "BaseNamG" 'Or "BASENAMG"
                    oEditColumn.TitleObject.Caption = "Transaction"
                    oEditColumn.Visible = bVisible
                Case "OrigNamG" 'Or "ORIGNAMG"
                    oEditColumn.TitleObject.Caption = "Orig. Transaction"
                    oEditColumn.Visible = bVisible
                Case Else
                    oEditColumn.TitleObject.Caption = sColumnUID
                    oEditColumn.Visible = bVisible
            End Select
        Next
    End Sub
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
    'Friend Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
    '    Dim BubbleEvent As Boolean = True
    '    Try
    '        If pVal.Before_Action = True Then
    '            Select Case pVal.ItemUID
    '                Case "mxGrid"
    '                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
    '                        MsgBox(pVal.ColUID)
    '                        Select Case pVal.ColUID
    '                            Case "BaseNam", "BaseNam|V"
    '                                Dim iRow As Integer = pVal.Row
    '                                Dim sTempEntry As String = String.Empty
    '                                Dim sTempType As String = String.Empty
    '                                iRow = oGridCl.GetDataTableRowIndex(pVal.Row)

    '                                sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("BaseEnt", iRow)
    '                                sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("BaseType", iRow)

    '                                If (sTempType = "0") Then
    '                                    Return False
    '                                End If
    '                                sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

    '                                oLnObj.LinkedObjectType = sTempType
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
    '                                oLnItm.Click(SAPbouiCOM.BoCellClickType.ct_Linked)
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
    '                                Return False

    '                            Case "TransNam", "TransNam|V"
    '                                Dim iRow As Integer = pVal.Row
    '                                Dim sTempEntry As String = String.Empty
    '                                Dim sTempType As String = String.Empty
    '                                iRow = oGridCl.GetDataTableRowIndex(pVal.Row)
    '                                sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("TransEnt", iRow)
    '                                sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("TransType", iRow)
    '                                If (sTempType = "0") Then
    '                                    Return False
    '                                End If
    '                                sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

    '                                oLnObj.LinkedObjectType = sTempType
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
    '                                oLnItm.Click(SAPbouiCOM.BoCellClickType.ct_Linked)
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
    '                                Return False
    '                            Case "OrigNam" Or "OrigNam|V"
    '                                Dim iRow As Integer = pVal.Row
    '                                Dim sTempEntry As String = String.Empty
    '                                Dim sTempType As String = String.Empty
    '                                iRow = oGridCl.GetDataTableRowIndex(pVal.Row)
    '                                sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("OrigEnt", iRow)
    '                                sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("OrigType", iRow)
    '                                If (sTempType = "0") Then
    '                                    Return False
    '                                End If
    '                                sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

    '                                oLnObj.LinkedObjectType = sTempType
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
    '                                oLnItm.Click(SAPbouiCOM.BoCellClickType.ct_Linked)
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
    '                                Return False
    '                            Case "RefOutNam" ''"RefOutNam|V"
    '                                Dim iRow As Integer = pVal.Row
    '                                Dim sTempEntry As String = String.Empty
    '                                Dim sTempType As String = String.Empty
    '                                iRow = oGridCl.GetDataTableRowIndex(pVal.Row)
    '                                sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("RefOutEnt", iRow)
    '                                sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("RefOutType", iRow)
    '                                If (sTempType = "0") Then
    '                                    Return False
    '                                End If
    '                                sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

    '                                oLnObj.LinkedObjectType = sTempType
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
    '                                oLnItm.Click(SAPbouiCOM.BoCellClickType.ct_Linked)
    '                                oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
    '                                Return False
    '                        End Select
    '                    End If
    '                Case "2"
    '                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
    '                        If ((Not (statusThread Is Nothing)) AndAlso ((statusThread.ThreadState = ThreadState.Running) Or (statusThread.ThreadState = ThreadState.WaitSleepJoin))) Then
    '                            statusThread.Suspend()
    '                            ' statusThread.Abort()
    '                        End If
    '                        If ((Not (myThread Is Nothing)) AndAlso ((myThread.ThreadState = ThreadState.Running) Or (myThread.ThreadState = ThreadState.WaitSleepJoin))) Then
    '                            myThread.Suspend()
    '                            '   myThread.Abort()
    '                        End If
    '                    End If

    '                Case "btPrint"
    '                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
    '                        oFormFIFO1.Close()
    '                        'BubbleEvent = ValidateParameters()
    '                    End If
    '            End Select

    '            Select Case pVal.EventType
    '                Case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD, SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
    '                    If ((Not (statusThread Is Nothing)) AndAlso ((statusThread.ThreadState = ThreadState.Running) Or (statusThread.ThreadState = ThreadState.WaitSleepJoin))) Then
    '                        statusThread.Suspend()
    '                        'statusThread.Abort()
    '                    End If
    '                    If ((Not (myThread Is Nothing)) AndAlso ((myThread.ThreadState = ThreadState.Running) Or (myThread.ThreadState = ThreadState.WaitSleepJoin))) Then
    '                        myThread.Suspend()
    '                        ' myThread.Abort()
    '                    End If
    '                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
    '                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
    '                    Dim sCFL_ID As String = String.Empty

    '                    oCFLEvento = pVal
    '                    sCFL_ID = oCFLEvento.ChooseFromListUID
    '                    Dim myVal As String = String.Empty
    '                    Dim compareVal As String = String.Empty
    '                    Select Case pVal.ItemUID
    '                        Case "tbItmGrpFr"
    '                            myVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
    '                            compareVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx
    '                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
    '                        Case "tbItmGrpTo"
    '                            myVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx
    '                            compareVal = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
    '                            Return setChooseFromListConditions(myVal, compareVal, sCFL_ID)
    '                    End Select
    '            End Select

    '        Else 'before action = false
    '            Select Case pVal.EventType
    '                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
    '                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
    '                    Dim sCFL_ID, sItemCode, sWareCode As String
    '                    Dim oCFL As SAPbouiCOM.ChooseFromList
    '                    Dim oDataTable As SAPbouiCOM.DataTable
    '                    Dim sItemGrpName As String = ""
    '                    Dim sItemGrpCod As String = ""

    '                    Select Case pVal.ItemUID
    '                        Case "tbItemFr", "tbItemTo"
    '                            oCFLEvento = pVal
    '                            sCFL_ID = oCFLEvento.ChooseFromListUID
    '                            oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
    '                            oDataTable = oCFLEvento.SelectedObjects
    '                            Try
    '                                sItemCode = oDataTable.GetValue("ItemCode", 0)
    '                            Catch ex As Exception

    '                            End Try
    '                            Select Case pVal.ItemUID
    '                                Case "tbItemFr"
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("uItemFr").ValueEx = sItemCode
    '                                Case "tbItemTo"
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("uItemTo").ValueEx = sItemCode
    '                            End Select
    '                        Case "tbItmGrpFr", "tbItmGrpTo"
    '                            oCFLEvento = pVal
    '                            sCFL_ID = oCFLEvento.ChooseFromListUID
    '                            oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
    '                            oDataTable = oCFLEvento.SelectedObjects

    '                            Try
    '                                sItemGrpName = oDataTable.GetValue("ItmsGrpNam", 0)
    '                                sItemGrpCod = oDataTable.GetValue("ItmsGrpCod", 0)
    '                            Catch ex As Exception

    '                            End Try
    '                            Select Case pVal.ItemUID
    '                                Case "tbItmGrpFr"
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx = sItemGrpName
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("tbItmGFr").ValueEx = sItemGrpCod

    '                                Case "tbItmGrpTo"
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx = sItemGrpName
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("tbItmGTo").ValueEx = sItemGrpCod
    '                            End Select
    '                        Case "tbWareFr", "tbWareTo"
    '                            oCFLEvento = pVal
    '                            sCFL_ID = oCFLEvento.ChooseFromListUID
    '                            oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
    '                            oDataTable = oCFLEvento.SelectedObjects

    '                            Try
    '                                sWareCode = oDataTable.GetValue("WhsCode", 0)
    '                            Catch ex As Exception

    '                            End Try
    '                            Select Case pVal.ItemUID
    '                                Case "tbWareFr"
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("uWareFr").ValueEx = sWareCode
    '                                Case "tbWareTo"
    '                                    oFormFIFO1.DataSources.UserDataSources.Item("uWareTo").ValueEx = sWareCode
    '                            End Select
    '                    End Select
    '            End Select
    '        End If
    '    Catch ex As Exception
    '        BubbleEvent = False
    '        MsgBox("ItemVeent : " & ex.ToString)
    '        SBO_Application.StatusBar.SetText("[ItemEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '    End Try
    '    Return BubbleEvent
    'End Function
    Friend Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action Then
                Select Case pVal.ItemUID
                    Case "mxGrid"
                        Select Case pVal.ColUID
                            Case "BaseNam", "BaseNam|V"
                                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                                    Dim iRow As Integer = pVal.Row
                                    Dim sTempEntry As String = String.Empty
                                    Dim sTempType As String = String.Empty
                                    Dim oFormFIFO1 As SAPbouiCOM.Form = SBO_Application.Forms.Item(FRM_NCM_SES2)
                                    Dim oGrid As SAPbouiCOM.Grid = oFormFIFO1.Items.Item("mxGrid").Specific
                                    Dim oLink As SAPbouiCOM.LinkedButton = oFormFIFO1.Items.Item("lbObj").Specific

                                    iRow = oGrid.GetDataTableRowIndex(pVal.Row)
                                    sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("BASEENT", iRow)
                                    sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("BASETYPE", iRow)
                                    If (sTempType = "0") Then
                                        Return False
                                    End If
                                    sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

                                    oLink.LinkedObjectType = sTempType
                                    oFormFIFO1.Items.Item("lbObj").LinkTo = "txtLink"
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
                                    oFormFIFO1.Items.Item("lbObj").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
                                    Return False
                                End If
                            Case "TransNam", "TransNam|V"
                                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                                    Dim iRow As Integer = pVal.Row
                                    Dim sTempEntry As String = String.Empty
                                    Dim sTempType As String = String.Empty
                                    Dim oFormFIFO1 As SAPbouiCOM.Form = SBO_Application.Forms.Item(FRM_NCM_SES2)
                                    Dim oGrid As SAPbouiCOM.Grid = oFormFIFO1.Items.Item("mxGrid").Specific
                                    Dim oLink As SAPbouiCOM.LinkedButton = oFormFIFO1.Items.Item("lbObj").Specific

                                    iRow = oGrid.GetDataTableRowIndex(pVal.Row)
                                    sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("TRANSENT", iRow)
                                    sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("TRANSTYPE", iRow)
                                    If (sTempType = "0") Then
                                        Return False
                                    End If
                                    sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

                                    oLink.LinkedObjectType = sTempType
                                    oFormFIFO1.Items.Item("lbObj").LinkTo = "txtLink"
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
                                    oFormFIFO1.Items.Item("lbObj").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
                                    Return False
                                End If
                            Case "OrigNam", "OrigNam|V"
                                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                                    Dim iRow As Integer = pVal.Row
                                    Dim sTempEntry As String = String.Empty
                                    Dim sTempType As String = String.Empty
                                    Dim oFormFIFO1 As SAPbouiCOM.Form = SBO_Application.Forms.Item(FRM_NCM_SES2)
                                    Dim oGrid As SAPbouiCOM.Grid = oFormFIFO1.Items.Item("mxGrid").Specific
                                    Dim oLink As SAPbouiCOM.LinkedButton = oFormFIFO1.Items.Item("lbObj").Specific

                                    iRow = oGrid.GetDataTableRowIndex(pVal.Row)
                                    sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("ORIGENT", iRow)
                                    sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("ORIGTYPE", iRow)
                                    If (sTempType = "0") Then
                                        Return False
                                    End If
                                    sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

                                    oLink.LinkedObjectType = sTempType
                                    oFormFIFO1.Items.Item("lbObj").LinkTo = "txtLink"
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
                                    oFormFIFO1.Items.Item("lbObj").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
                                    Return False
                                End If
                            Case "RefOutNam", "RefOutNam|V"
                                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED) Then
                                    Dim iRow As Integer = pVal.Row
                                    Dim sTempEntry As String = String.Empty
                                    Dim sTempType As String = String.Empty
                                    Dim oFormFIFO1 As SAPbouiCOM.Form = SBO_Application.Forms.Item(FRM_NCM_SES2)
                                    Dim oGrid As SAPbouiCOM.Grid = oFormFIFO1.Items.Item("mxGrid").Specific
                                    Dim oLink As SAPbouiCOM.LinkedButton = oFormFIFO1.Items.Item("lbObj").Specific

                                    iRow = oGrid.GetDataTableRowIndex(pVal.Row)
                                    sTempEntry = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("REFOUTENT", iRow)
                                    sTempType = oFormFIFO1.DataSources.DataTables.Item("abc").GetValue("REFOUTTYPE", iRow)
                                    If (sTempType = "0") Then
                                        Return False
                                    End If
                                    sTempType = IIf(String.Compare(sTempType, "58", True) = 0, "30", sTempType)

                                    oLink.LinkedObjectType = sTempType
                                    oFormFIFO1.Items.Item("lbObj").LinkTo = "txtLink"
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = sTempEntry
                                    oFormFIFO1.Items.Item("lbObj").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    oFormFIFO1.DataSources.UserDataSources.Item("txtLink").ValueEx = String.Empty
                                    Return False
                                End If
                        End Select
                End Select
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "btPrint" Then
                            Dim oFormFIFO1 As SAPbouiCOM.Form = SBO_Application.Forms.Item(FRM_NCM_SES2)
                            oFormFIFO1.Close()
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
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Select Case pVal.ItemUID
                            Case "tbItemFr", "tbItemTo"
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim sCFL_ID, sItemCode As String
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
                                Dim sItemGrpName As String = ""
                                Dim sItemGrpCod As String = ""
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
                                Dim sCFL_ID, sWareCode As String
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
