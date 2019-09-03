Public Class MRPDemandandSupplyRpt
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox

    Private oConditions As SAPbouiCOM.Conditions
    Private oCondition As SAPbouiCOM.Condition
    Private ds As DataSet
    Private dtRpt As System.Data.DataTable
    Private dtGeneral As System.Data.DataTable
    Private file_MRPDemandSupplyRpt As String = "ncmMRPDemandandSupplyRpt.srf"
    Private frm_MRPDemandSupplyRpt As String = "ncmMRPDemandSupplyReport"
    Private g_sReportFileName As String = ""

    Public Sub New()
        MyBase.new()
    End Sub

#Region "General Functions"
    Public Sub ShowForm()
        If LoadFromXML(file_MRPDemandSupplyRpt) = True Then
            Try
                oForm = SBO_Application.Forms.Item(frm_MRPDemandSupplyRpt)
                ds = New dsMRP_RPT

                oForm.DataSources.UserDataSources.Add("ItemFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                oForm.DataSources.UserDataSources.Add("ItemTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                oForm.DataSources.UserDataSources.Add("MinInv", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

                oEdit = oForm.Items.Item("txtItemFr").Specific
                oEdit.DataBind.SetBound(True, "", "ItemFr")
                oEdit = oForm.Items.Item("txtItemTo").Specific
                oEdit.DataBind.SetBound(True, "", "ItemTo")
                oCheck = oForm.Items.Item("chkMinInv").Specific
                oCheck.DataBind.SetBound(True, "", "MinInv")
                oCheck.ValOff = "N"
                oCheck.ValOn = "Y"

                oCombo = oForm.Items.Item("cbItmGrpFr").Specific
                PopulateItemGroupCombo(oCombo)
                oCombo = oForm.Items.Item("cbItmGrpTo").Specific
                PopulateItemGroupCombo(oCombo)
                oCombo = oForm.Items.Item("cbFct").Specific
                PopulateForeCastCombo(oCombo)

                SetChooseFromList()
                oForm.Visible = True
            Catch ex As Exception
                oForm.Update()
                oForm.Refresh()
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        Else
            Try
                oForm = SBO_Application.Forms.Item(frm_MRPDemandSupplyRpt)
                If oForm.Visible Then
                    oForm.Select()
                Else
                    oForm.Close()
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Function SetChooseFromList() As Boolean
        Try
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            oCFLs = oForm.ChooseFromLists
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "4"
            oCFLCreationParams.UniqueID = "CFL_ITEMFR"
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCFLCreationParams.UniqueID = "CFL_ITEMTO"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oEdit = oForm.Items.Item("txtItemFr").Specific
            oEdit.ChooseFromListUID = "CFL_ITEMFR"
            oEdit.ChooseFromListAlias = "ItemCode"
            oEdit = oForm.Items.Item("txtItemTo").Specific
            oEdit.ChooseFromListUID = "CFL_ITEMTO"
            oEdit.ChooseFromListAlias = "ItemCode"

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[SetChooseFromList]:" & ex.Message)
            Return False
        End Try
    End Function
    Friend Sub PopulateItemGroupCombo(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oCombo.ValidValues.Count > 0 Then
                For i As Integer = 1 To oCombo.ValidValues.Count
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            oRec.DoQuery("Select ItmsGrpCod, ItmsGrpNam From OITB Order by ItmsGrpCod")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                oCombo.ValidValues.Add("", "None")
                While Not oRec.EoF
                    oCombo.ValidValues.Add(oRec.Fields.Item("ItmsGrpCod").Value, oRec.Fields.Item("ItmsGrpNam").Value)
                    oRec.MoveNext()
                End While
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[PopulateItemGroupCombo]" & ex.Message)
        End Try
    End Sub
    Friend Sub PopulateForeCastCombo(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If oCombo.ValidValues.Count > 0 Then
                For i As Integer = 1 To oCombo.ValidValues.Count
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            oRec.DoQuery("Select AbsId, Name From OFCT Order By AbsId")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                oCombo.ValidValues.Add("", "None")
                While Not oRec.EoF
                    oCombo.ValidValues.Add(oRec.Fields.Item("AbsId").Value, oRec.Fields.Item("Name").Value)
                    oRec.MoveNext()
                End While
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[PopulateForeCastCombo]" & ex.Message)
        End Try
    End Sub
    Private Function Validate() As Boolean
        Dim ValueFrom As String = ""
        Dim ValueTo As String = ""
        Try

            'Must Fill ItemCode. For Interval
            oEdit = oForm.Items.Item("txtItemFr").Specific
            ValueFrom = oEdit.Value
            oEdit = oForm.Items.Item("txtItemTo").Specific
            ValueTo = oEdit.Value
            If (ValueFrom <> String.Empty And ValueTo = String.Empty) Or (ValueTo <> String.Empty And ValueFrom = String.Empty) Then
                Throw New Exception("Invalid Item Code. : You Must Fill Both Item Code")
            End If

            'Must Fill Item Group. For Interval
            oCombo = oForm.Items.Item("cbItmGrpFr").Specific
            ValueFrom = oCombo.Selected.Value
            oCombo = oForm.Items.Item("cbItmGrpTo").Specific
            ValueTo = oCombo.Selected.Value
            If (ValueFrom <> String.Empty And ValueTo = String.Empty) Or (ValueTo <> String.Empty And ValueFrom = String.Empty) Then
                Throw New Exception("Invalid Item Group. : You Must Fill Both Item Group")
            End If

            'Must Fill Item Code or Item Group
            oEdit = oForm.Items.Item("txtItemFr").Specific
            ValueFrom = oEdit.Value
            oCombo = oForm.Items.Item("cbItmGrpFr").Specific
            ValueTo = oCombo.Selected.Value
            If (ValueFrom = String.Empty And ValueTo = String.Empty) Then
                Throw New Exception("Invalid Parameter. : You Must Fill Item Code. Or Item Group.")
            End If
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[Validate]:" & ex.Message)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFileName = GetSharedFilePath(ReportName.MRPSupplyDemandReport)
            If g_sReportFileName <> "" Then
                If IsSharedFilePathExists(g_sReportFileName) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFileName = " "
            SBO_Application.StatusBar.SetText("[AP SOA].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Friend Sub LoadViewer()
        Dim frm As Hydac_FormViewer
        Dim bIsShared As Boolean = False
        Try
            dtRpt = ds.Tables("TableReport")
            dtGeneral = ds.Tables("General")
            dtRpt.Clear()
            dtGeneral.Clear()

            'Load report data to temp table
            bIsShared = IsSharedFileExist()
            SBO_Application.StatusBar.SetText("Prepare to load data...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            LoadPrintData()
            SBO_Application.StatusBar.SetText("Data loaded successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            frm = New Hydac_FormViewer
            frm.Report = ReportName.MRPSupplyDemandReport
            frm.IsShared = bIsShared
            frm.SharedReportName = g_sReportFileName
            frm.Dataset = ds
            frm.ShowDialog()
            SBO_Application.StatusBar.SetText("Operation ended successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Catch ex As Exception
            SBO_Application.MessageBox("[LoadViewer]:" & ex.Message)
        End Try
    End Sub
    Private Sub LoadPrintData()
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim ValueFrom, ValueTo As String
        Dim headerRow, GeneralRow As DataRow
        Try
            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Load For General
            GeneralRow = dtGeneral.NewRow
            GeneralRow("Company") = oCompany.CompanyName

            With oForm.DataSources.UserDataSources
                sQuery = "EXEC NCM_PROC_MRPDEMANDSUPPLY_RPT "
                If .Item("ItemFr").ValueEx <> "" Then
                    sQuery &= .Item("ItemFr").ValueEx & ", " & .Item("ItemTo").ValueEx & ", "
                Else
                    sQuery &= "'', '', "
                End If

                GeneralRow("ItemFrom") = .Item("ItemFr").ValueEx
                GeneralRow("ItemTo") = .Item("ItemTo").ValueEx

                oCombo = oForm.Items.Item("cbItmGrpFr").Specific
                ValueFrom = oCombo.Selected.Value
                GeneralRow("ItemGroupFrom") = oCombo.Selected.Description
                oCombo = oForm.Items.Item("cbItmGrpTo").Specific
                ValueTo = oCombo.Selected.Value
                GeneralRow("ItemGroupTo") = oCombo.Selected.Description

                If ValueFrom <> "" Then
                    sQuery &= ValueFrom & ", " & ValueTo & ", "
                Else
                    sQuery &= "'', '', "
                End If

                oCombo = oForm.Items.Item("cbFct").Specific
                ValueFrom = oCombo.Selected.Value
                GeneralRow("ForeCast") = oCombo.Selected.Description

                If ValueFrom <> "" Then
                    sQuery &= ValueFrom & ", "
                Else
                    sQuery &= "'', "
                End If

                If .Item("MinInv").ValueEx <> "" Then
                    sQuery &= .Item("MinInv").ValueEx & " "
                Else
                    sQuery &= "'N' "
                End If

                dtGeneral.Rows.Add(GeneralRow)

                oRecordset.DoQuery(sQuery)
                If oRecordset.RecordCount > 0 Then
                    For i As Integer = 1 To oRecordset.RecordCount
                        headerRow = dtRpt.NewRow
                        headerRow("ItemCode") = oRecordset.Fields.Item("ItemCode").Value
                        headerRow("ItemName") = oRecordset.Fields.Item("ItemName").Value
                        headerRow("MinStock") = oRecordset.Fields.Item("MinStock").Value
                        headerRow("OnHandQty") = oRecordset.Fields.Item("OnHandQty").Value
                        headerRow("SupplyQty") = oRecordset.Fields.Item("SupplyQty").Value
                        headerRow("DemandQty") = oRecordset.Fields.Item("DemandQty").Value
                        headerRow("OverDueSupply") = oRecordset.Fields.Item("OverDueSupply").Value
                        headerRow("OverDueDemand") = oRecordset.Fields.Item("OverDueDemand").Value
                        headerRow("First30Supply") = oRecordset.Fields.Item("First30Supply").Value
                        headerRow("First30Demand") = oRecordset.Fields.Item("First30Demand").Value
                        headerRow("Second30Supply") = oRecordset.Fields.Item("Second30Supply").Value
                        headerRow("Second30Demand") = oRecordset.Fields.Item("Second30Demand").Value
                        headerRow("Third30Supply") = oRecordset.Fields.Item("Third30Supply").Value
                        headerRow("Third30Demand") = oRecordset.Fields.Item("Third30Demand").Value
                        headerRow("Fourth30Supply") = oRecordset.Fields.Item("Fourth30Supply").Value
                        headerRow("Fourth30Demand") = oRecordset.Fields.Item("Fourth30Demand").Value
                        headerRow("Fifth30Supply") = oRecordset.Fields.Item("Fifth30Supply").Value
                        headerRow("Fifth30Demand") = oRecordset.Fields.Item("Fifth30Demand").Value
                        headerRow("Sixth30Supply") = oRecordset.Fields.Item("Sixth30Supply").Value
                        headerRow("Sixth30Demand") = oRecordset.Fields.Item("Sixth30Demand").Value
                        headerRow("Seventh30Supply") = oRecordset.Fields.Item("Seventh30Supply").Value
                        headerRow("Seventh30Demand") = oRecordset.Fields.Item("Seventh30Demand").Value
                        headerRow("Eight30Supply") = oRecordset.Fields.Item("Eight30Supply").Value
                        headerRow("Eight30Demand") = oRecordset.Fields.Item("Eight30Demand").Value
                        headerRow("Nineth30Supply") = oRecordset.Fields.Item("Nineth30Supply").Value
                        headerRow("Nineth30Demand") = oRecordset.Fields.Item("Nineth30Demand").Value
                        headerRow("Ten30Supply") = oRecordset.Fields.Item("Ten30Supply").Value
                        headerRow("Ten30Demand") = oRecordset.Fields.Item("Ten30Demand").Value
                        headerRow("Eleven30Supply") = oRecordset.Fields.Item("Eleven30Supply").Value
                        headerRow("Eleven30Demand") = oRecordset.Fields.Item("Eleven30Demand").Value
                        headerRow("Twelve30Supply") = oRecordset.Fields.Item("Twelve30Supply").Value
                        headerRow("Twelve30Demand") = oRecordset.Fields.Item("Twelve30Demand").Value
                        headerRow("OverTwelveSupply") = oRecordset.Fields.Item("OverTwelveSupply").Value
                        headerRow("OverTwelveDemand") = oRecordset.Fields.Item("OverTwelveDemand").Value
                        dtRpt.Rows.Add(headerRow)
                        oRecordset.MoveNext()
                    Next
                End If
            End With
        Catch ex As Exception
            SBO_Application.MessageBox("[LoadPrintData]:" & ex.Message)
        End Try
    End Sub
#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.Before_Action = True Then
            Else
                If pval.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oCFLEvento = pval

                    Dim sCode As String = ""
                    oDataTable = oCFLEvento.SelectedObjects
                    If Not (oDataTable Is Nothing) Then
                        If oCFLEvento.ChooseFromListUID = "CFL_ITEMFR" Then
                            sCode = oDataTable.GetValue("ItemCode", 0)
                            oForm.DataSources.UserDataSources.Item("ItemFr").ValueEx = sCode
                        End If
                        If oCFLEvento.ChooseFromListUID = "CFL_ITEMTO" Then
                            sCode = oDataTable.GetValue("ItemCode", 0)
                            oForm.DataSources.UserDataSources.Item("ItemTo").ValueEx = sCode
                        End If
                    End If
                End If

                If pval.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                    If pval.ItemUID = "btnOK" Then
                        If Validate() Then
                            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                            myThread.SetApartmentState(Threading.ApartmentState.STA)
                            myThread.Start()
                        Else
                            BubbleEvent = False
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
