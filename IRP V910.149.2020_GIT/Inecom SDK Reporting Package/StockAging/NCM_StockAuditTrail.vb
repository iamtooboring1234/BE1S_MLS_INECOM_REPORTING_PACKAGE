
Option Strict Off
Option Explicit On 

Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class NCM_StockAuditTrail

#Region "Global Variables"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oFormFIFO1 As SAPbouiCOM.Form
    Private sqlConn As SqlConnection

    Private sErrMsg As String
    Private lErrCode As Integer
    Private da As SqlDataAdapter
    Private dr As SqlDataReader
    Private ds As DataSet

    Private dtFIFO As DataTable
    Private dtExportD As DataTable
    Private dtExportS As DataTable
    Dim bIsExportToExcel As Boolean = False
    Dim sRptType As String = String.Empty

    Private g_sReportFilename As String = ""
    Private g_iSecond As Integer
    Dim aTimer As New System.Windows.Forms.Timer
    Private g_bIsShared As Boolean = False
    Private g_bIsXSDShared As Boolean = False
    Dim statusThread As System.Threading.Thread
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
        AddHandler aTimer.Tick, AddressOf Timer_Tick
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region

#Region "Populate Data"
    Friend Sub LoadForm()
        Try
            Dim oEdit As SAPbouiCOM.EditText
            Dim oLink As SAPbouiCOM.LinkedButton
            Dim sCode1 As String = String.Empty
            Dim sCode2 As String = String.Empty
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            If LoadFromXML("Inecom_SDK_Reporting_Package." & FRM_NCM_SES1 & ".srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading Form...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_SES1)

                AddChooseFromList()
                oStaticLn = DirectCast(oFormFIFO1.Items.Item("lbTimer").Specific, SAPbouiCOM.StaticText)

                With oFormFIFO1.DataSources.UserDataSources
                    .Add("uItemFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uItemTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uWareFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uWareTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("uAsDate", SAPbouiCOM.BoDataType.dt_DATE)
                    .Add("tbItmGrpFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                    .Add("tbItmGrpTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                End With

                '' Bind Form Controls
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
                oFormFIFO1.DataSources.UserDataSources.Item("uAsDate").ValueEx = DateTime.Now.ToString("yyyyMMdd")

                SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                oFormFIFO1.Visible = True
            Else
                Try
                    oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_SES1)
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
            SBO_Application.MessageBox("[AddChooseFromList] : " & vbNewLine & ex.Message)
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
    Private Sub Timer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        g_iSecond += 1
        SBO_Application.StatusBar.SetText("Processing " & g_iSecond & " seconds ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Private Function ExecuteProcedure() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""
            With oFormFIFO1.DataSources.UserDataSources
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
            oRec.DoQuery(sQuery)
            oRec = Nothing

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[ExecProd] : " & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Private Function GenerateRecords() As Boolean
        Try
            Dim sQuery As String = ""
            Dim oRec As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim inputString As String() = New String(12) {}
            With oFormFIFO1.DataSources.UserDataSources
                inputString(0) = .Item("uAsDate").ValueEx
                inputString(1) = oCompany.UserSignature.ToString()
                inputString(2) = .Item("uItemFr").ValueEx.Replace("'", "''")
                inputString(3) = .Item("uItemTo").ValueEx.Replace("'", "''")
                inputString(4) = .Item("uWareFr").ValueEx.Replace("'", "''")
                inputString(5) = .Item("uWareTo").ValueEx.Replace("'", "''")
                inputString(6) = .Item("tbItmGrpFr").ValueEx.Replace("'", "''")
                inputString(7) = .Item("tbItmGrpTo").ValueEx.Replace("'", "''")
            End With
            ' '' ----------------------------------------------------------------------
            'sQuery = "SELECT ""U_QUERY"" FROM ""@NCM_QUERY"" WHERE ""U_TYPE"" ='NCM_SES_MOV1'"
            'oRec.DoQuery(sQuery)
            'If oRec.RecordCount > 0 Then
            '    oRec.MoveFirst()
            '    sQuery = oRec.Fields.Item(0).Value
            '    sQuery = String.Format(sQuery, inputString)            
            '    '' ----------------------------------------------------------------------
            'End If
            sQuery = "CALL NCM_SP_SES_MOV1 ("
            sQuery &= "'" & inputString(0) & "',"
            sQuery &= "'" & inputString(1) & "',"
            sQuery &= "'" & inputString(2) & "',"
            sQuery &= "'" & inputString(3) & "',"
            sQuery &= "'" & inputString(4) & "',"
            sQuery &= "'" & inputString(5) & "',"
            sQuery &= "'" & inputString(6) & "',"
            sQuery &= "'" & inputString(7) & "')"
            oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            oRec = Nothing
            Return True

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GenerateRecords] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub PrintStatus()
        Dim i As Integer = 0
        Do While (1 = 1)
            Select Case i
                Case 0
                    oStaticLn.Caption = "Printing Report"
                Case 1
                    oStaticLn.Caption = "Printing Report."
                Case 2
                    oStaticLn.Caption = "Printing Report.."
                Case 3
                    oStaticLn.Caption = "Printing Report..."
                Case 4
                    oStaticLn.Caption = "Printing Report...."
            End Select
            'oStaticLn.Caption = "Printing Report"
            System.Threading.Thread.CurrentThread.Sleep(1000)
            If (i = 4) Then
                i = 0
            Else
                i = i + 1
            End If
        Loop
    End Sub
    Private Sub PrintSAR_FIFO_NonBatch()
        Try
            statusThread = New System.Threading.Thread(AddressOf PrintStatus)
            Try
                oFormFIFO1.Items.Item("btPrint").Enabled = False
                If ExecuteProcedure() Then
                    If GenerateRecords() Then
                        oStockAudit_Out.LoadForm()
                    End If
                End If
            Catch ex As Exception
                SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                aTimer.Enabled = False
                oFormFIFO1.Items.Item("btPrint").Enabled = True
                g_iSecond = 0
            End Try
        Catch ex As Exception

        Finally
            If ((Not (statusThread Is Nothing)) AndAlso ((statusThread.ThreadState = ThreadState.Running) Or (statusThread.ThreadState = ThreadState.WaitSleepJoin))) Then
                statusThread.Suspend()
                'statusThread.Abort()
            End If
            System.Threading.Thread.CurrentThread.Sleep(500)
            oStaticLn.Caption = String.Empty
        End Try
    End Sub
    Private Function ValidateParameters() As Boolean
        oFormFIFO1.ActiveItem = "tbItemFr"
        Dim sFromValue As String = ""
        Dim sToValue As String = ""

        Try
            With oFormFIFO1.DataSources.UserDataSources
                sFromValue = .Item("uItemFr").ValueEx
                sToValue = .Item("uItemTo").ValueEx
                If (Not (sFromValue.Length = 0 AndAlso sToValue.Length = 0)) Then
                    If (String.Compare(sFromValue, sToValue, True) > 0) Then
                        oFormFIFO1.ActiveItem = "tbItemFr"
                        SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - item from is greater than item to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
                sFromValue = .Item("uWareFr").ValueEx
                sToValue = .Item("uWareTo").ValueEx
                If (Not (sFromValue.Length = 0 AndAlso sToValue.Length = 0)) Then
                    If (String.Compare(sFromValue, sToValue, True) > 0) Then
                        oFormFIFO1.ActiveItem = "tbWareFr"
                        SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - warehouse from is greater than warehouse to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
                sFromValue = .Item("tbItmGrpFr").ValueEx
                sToValue = .Item("tbItmGrpTo").ValueEx
                If (Not (sFromValue.Length = 0 AndAlso sToValue.Length = 0)) Then
                    If (String.Compare(sFromValue, sToValue, True) > 0) Then
                        oFormFIFO1.ActiveItem = "tbItmGrpFr"
                        SBO_Application.StatusBar.SetText("[StockAging][ValidateParameters] - item group from is greater than item group to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            End With

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
                        If (pVal.ItemUID = "2") Then
                            If ((Not (statusThread Is Nothing)) AndAlso ((statusThread.ThreadState = ThreadState.Running) Or (statusThread.ThreadState = ThreadState.WaitSleepJoin))) Then
                                statusThread.Suspend()
                                'statusThread.Abort()
                            End If
                            If ((Not (myThread Is Nothing)) AndAlso ((myThread.ThreadState = ThreadState.Running) Or (myThread.ThreadState = ThreadState.WaitSleepJoin))) Then
                                myThread.Suspend()
                                'myThread.Abort()
                            End If
                        ElseIf pVal.ItemUID = "btPrint" Then
                            BubbleEvent = ValidateParameters()
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD, SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        If ((Not (statusThread Is Nothing)) AndAlso ((statusThread.ThreadState = ThreadState.Running) Or (statusThread.ThreadState = ThreadState.WaitSleepJoin))) Then
                            statusThread.Suspend()
                            ' statusThread.Abort()
                        End If
                        If ((Not (myThread Is Nothing)) AndAlso ((myThread.ThreadState = ThreadState.Running) Or (myThread.ThreadState = ThreadState.WaitSleepJoin))) Then
                            myThread.Suspend()
                            ' myThread.Abort()
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
                                Dim sItemGrpName As String = ""
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oDataTable As SAPbouiCOM.DataTable

                                oCFLEvento = pVal
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                oCFL = oFormFIFO1.ChooseFromLists.Item(sCFL_ID)
                                oDataTable = oCFLEvento.SelectedObjects

                                Try
                                    sItemGrpName = oDataTable.GetValue("ItmsGrpNam", 0)
                                Catch ex As Exception

                                End Try
                                Select Case pVal.ItemUID
                                    Case "tbItmGrpFr"
                                        oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx = sItemGrpName
                                    Case "tbItmGrpTo"
                                        oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx = sItemGrpName
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
