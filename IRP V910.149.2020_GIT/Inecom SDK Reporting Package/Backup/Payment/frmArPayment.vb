Imports System.Threading
Imports SAPbobsCOM

Public Class frmARPayment
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private dsARPayment As System.Data.DataSet
    Private sBPFrom As String = ""
    Private sBPTo As String = ""
    Private sDateFrom As String = ""
    Private sDateTo As String = ""
    Private g_sReportFileName As String = ""
    Const C_FRM_AR_PAYMENT As String = "NCM_AR_PAYMENT"

#Region "General Functions"
    Public Sub LoadForm()
        Try
            'Close Existing Form if any
            For i As Integer = 0 To SBO_Application.Forms.Count - 1
                If SBO_Application.Forms.Item(i).TypeEx = C_FRM_AR_PAYMENT Then
                    SBO_Application.Forms.Item(i).Close()
                    Exit For
                End If
            Next
            If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_AR_PAYMENT.srf") Then
                oForm = SBO_Application.Forms.Item(C_FRM_AR_PAYMENT)
                oForm.SupportedModes = -1
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                AddChooseFromList()
                DefineUserDataSource()

                oForm.Items.Item("txtDateFr").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                oForm.Visible = True
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[Load_Form] : " & oForm.UniqueID & " : " & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddChooseFromList()
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = oForm.ChooseFromLists
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            'Get Card Code From
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oCFLCreationParams.UniqueID = "CFL_BP_Fr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL_Cust
            oCons = oCFL.GetConditions
            oCon = oCons.Add
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            'Get Card Code To
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oCFLCreationParams.UniqueID = "CFL_BP_To"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL_Cust
            oCons = oCFL.GetConditions
            oCon = oCons.Add
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AddChooseFromList] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub DefineUserDataSource()
        Try
            With oForm.DataSources.UserDataSources
                .Add("txtBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtDateFr", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("txtDateTo", SAPbouiCOM.BoDataType.dt_DATE)
            End With
            oEdit = oForm.Items.Item("txtBPFr").Specific
            oEdit.DataBind.SetBound(True, , "txtBPFr")
            oEdit.ChooseFromListUID = "CFL_BP_Fr"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = oForm.Items.Item("txtBPTo").Specific
            oEdit.DataBind.SetBound(True, , "txtBPTo")
            oEdit.ChooseFromListUID = "CFL_BP_To"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = oForm.Items.Item("txtDateFr").Specific
            oEdit.DataBind.SetBound(True, , "txtDateFr")
            oEdit = oForm.Items.Item("txtDateTo").Specific
            oEdit.DataBind.SetBound(True, , "txtDateTo")
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[Define UserDataSource] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function PrepareDataset() As Boolean
        Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
        Dim sQuery As String = ""
        Dim sPaymentNo As String = ""
        Dim objDataRow As System.Data.DataRow
        Dim dt As System.Data.DataTable

        Try
            dsARPayment = New dsPayment
            dt = dsARPayment.Tables("dtPayment")
            dt.Clear()
            With oForm.DataSources.UserDataSources
                sBPFrom = Convert.ToString(.Item("txtBPFr").ValueEx).Trim
                sBPTo = Convert.ToString(.Item("txtBPTo").ValueEx).Trim
                sDateFrom = Convert.ToString(.Item("txtDateFr").ValueEx).Trim
                sDateTo = Convert.ToString(.Item("txtDateTo").ValueEx).Trim
            End With

            sQuery = "EXEC NCM_RPT_PAYMENT '" & sDateFrom & "','" & sDateTo & "','" & sBPFrom & "','" & sBPTo & "','AR'"
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                While Not oRecord.EoF
                    sPaymentNo = oRecord.Fields.Item("PaymentNo").Value
                    SBO_Application.StatusBar.SetText("Loading data [" & sPaymentNo & "] ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    objDataRow = dt.NewRow

                    objDataRow("PaymentNo") = oRecord.Fields.Item("PaymentNo").Value
                    objDataRow("CardCode") = oRecord.Fields.Item("PaymentBPCode").Value
                    objDataRow("CardName") = oRecord.Fields.Item("PaymentBPName").Value
                    objDataRow("DocDate") = IIf(oRecord.Fields.Item("PaymentDate").Value = GetDateObject("18991230"), System.DBNull.Value, oRecord.Fields.Item("PaymentDate").Value)
                    objDataRow("PaymentRef") = oRecord.Fields.Item("PaymentRef").Value
                    objDataRow("InvoiceType") = oRecord.Fields.Item("InvoiceType").Value

                    objDataRow("InvoiceNo") = oRecord.Fields.Item("InvoiceNo").Value
                    objDataRow("BPRefNo") = oRecord.Fields.Item("InvoiceBPRefNo").Value
                    objDataRow("InvoiceAmtFC") = oRecord.Fields.Item("InvoiceAmtFC").Value
                    objDataRow("InvoiceAmtLC") = oRecord.Fields.Item("InvoiceAmtLC").Value
                    objDataRow("InvoiceCurrency") = oRecord.Fields.Item("InvoiceCurr").Value
                    objDataRow("CheckNo") = oRecord.Fields.Item("CheckNo").Value
                    objDataRow("TransferNo") = oRecord.Fields.Item("TransferNo").Value
                    objDataRow("PaymentSumLC") = oRecord.Fields.Item("PaymentSumLC").Value
                    objDataRow("PaymentSumFC") = oRecord.Fields.Item("PaymentSumFC").Value
                    objDataRow("PaymentCurr") = oRecord.Fields.Item("PaymentCurr").Value
                    objDataRow("PaymentRate") = oRecord.Fields.Item("PaymentRate").Value
                    dt.Rows.Add(objDataRow)
                    oRecord.MoveNext()
                End While
            End If

            oRecord = Nothing
            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PrepareDataset]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFileName = GetSharedFilePath(ReportName.ARPayment)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[AR Payment].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        Try
            Dim frm As New Hydac_FormViewer
            If PrepareDataset() Then
                frm.IsShared = IsSharedFileExist()
                frm.SharedReportName = g_sReportFileName
                frm.ReportName = ReportName.ARPayment
                frm.ReportPath = g_sReportFileName
                frm.ReportDataSet = dsARPayment
                frm.UserCode = oCompany.UserName
                frm.ParamDateFrom = IIf(sDateFrom <> "", Format(GetDateObject(sDateFrom), "dd.MM.yyyy"), "")
                frm.ParamDateTo = IIf(sDateTo <> "", Format(GetDateObject(sDateTo), "dd.MM.yyyy"), "")
                frm.ParamBPFrom = sBPFrom
                frm.ParamBPTo = sBPTo
                frm.ShowDialog()
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.Before_Action = True Then

            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pval.ItemUID = "btnView" Then
                            Dim MyThread As Thread = New Thread(New ThreadStart(AddressOf LoadViewer))
                            MyThread.SetApartmentState(ApartmentState.STA)
                            MyThread.Start()
                        End If
                    Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                        Dim sCFL_ID As String
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim oDataTable As SAPbouiCOM.DataTable

                        oCFLEvento = pval
                        sCFL_ID = oCFLEvento.ChooseFromListUID
                        oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                        oDataTable = oCFLEvento.SelectedObjects

                        If pval.ItemUID = "txtBPFr" Then
                            Try
                                oForm.DataSources.UserDataSources.Item(pval.ItemUID).ValueEx = oDataTable.GetValue("CardCode", 0)
                            Catch ex As Exception
                            End Try
                        ElseIf pval.ItemUID = "txtBPTo" Then
                            Try
                                oForm.DataSources.UserDataSources.Item(pval.ItemUID).ValueEx = oDataTable.GetValue("CardCode", 0)
                            Catch ex As Exception
                            End Try
                        End If
                End Select
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ItemEvent] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region
End Class
