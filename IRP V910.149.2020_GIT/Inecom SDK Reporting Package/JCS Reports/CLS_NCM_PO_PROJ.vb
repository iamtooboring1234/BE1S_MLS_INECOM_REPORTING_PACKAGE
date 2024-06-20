Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading
Imports System.IO

Public Class CLS_NCM_PO_PROJ

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oChck As SAPbouiCOM.CheckBox
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Private dtCommand As System.Data.DataTable
    Private ds As System.Data.DataSet
    Private sqlConn As SqlConnection
    Private sqlComm As SqlCommand
    Private da As SqlDataAdapter
    Private dtRpt As System.Data.DataTable
    Private dtGeneral As System.Data.DataTable
#End Region

#Region "Initialisation"
    Public Sub LoadForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_PO_PROJ.srf") Then
            oForm = SBO_Application.Forms.Item("NCM_PO_PROJ")
            ds = New dsPODetailByVendor
            AddDataSource()
            If (Not oForm.Visible) Then
                oForm.Visible = True
            End If
        Else
            Try
                oForm = SBO_Application.Forms.Item("NCM_PO_PROJ")
                If oForm.Visible = False Then
                    oForm.Close()
                Else
                    oForm.Select()
                End If
            Catch ex As Exception
            End Try
        End If
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

            ' Adding 3 CFLs, 1 for Supplier and 2 for User
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner

            oCFLCreationParams.UniqueID = "CFL_BPFr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            oCFLCreationParams.UniqueID = "CFL_BPTo"
            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ProjectCodes

            oCFLCreationParams.UniqueID = "CFL_PrjFr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.UniqueID = "CFL_PrjTo"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLs = oForm.ChooseFromLists
            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_ItemGroups

            oCFLCreationParams.UniqueID = "CFL_ItemGroupFr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCFLCreationParams.UniqueID = "CFL_ItemGroupTo"
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AddChooseFromList] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddDataSource()
        Try
            AddChooseFromList()
            With oForm.DataSources.UserDataSources
                .Add("tbDatFr", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("tbDatTo", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("tbPrjFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("tbPrjTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("tbBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("tbBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("ckDoc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("tbIGrpFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("tbIGrpTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            End With

            oEdit = oForm.Items.Item("tbDatFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbDatFr")
            oEdit = oForm.Items.Item("tbDatTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbDatTo")

            oEdit = oForm.Items.Item("tbBPFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbBPFr")
            oEdit.ChooseFromListUID = "CFL_BPFr"
            oEdit.ChooseFromListAlias = "CardCode"
            oEdit = oForm.Items.Item("tbBPTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbBPTo")
            oEdit.ChooseFromListUID = "CFL_BPTo"
            oEdit.ChooseFromListAlias = "CardCode"

            oEdit = oForm.Items.Item("tbPrjFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbPrjFr")
            oEdit.ChooseFromListUID = "CFL_PrjFr"
            oEdit.ChooseFromListAlias = "PrjCode"
            oEdit = oForm.Items.Item("tbPrjTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbPrjTo")
            oEdit.ChooseFromListUID = "CFL_PrjTo"
            oEdit.ChooseFromListAlias = "PrjCode"

            oEdit = oForm.Items.Item("tbIGrpFr").Specific
            oEdit.DataBind.SetBound(True, "", "tbIGrpFr")
            oEdit.ChooseFromListUID = "CFL_ItemGroupFr"
            oEdit.ChooseFromListAlias = "ItmsGrpCod"

            oEdit = oForm.Items.Item("tbIGrpTo").Specific
            oEdit.DataBind.SetBound(True, "", "tbIGrpTo")
            oEdit.ChooseFromListUID = "CFL_ItemGroupTo"
            oEdit.ChooseFromListAlias = "ItmsGrpCod"

            oChck = oForm.Items.Item("ckDoc").Specific
            oChck.DataBind.SetBound(True, "", "ckDoc")
            oChck.ValOff = "N"
            oChck.ValOn = "Y"
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AddDataSource] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub LoadViewer()
        Dim frm As Hydac_FormViewer
        Dim bIsShared As Boolean = False
        Try
            dtRpt = ds.Tables("TableReport")
            dtGeneral = ds.Tables("General")
            dtRpt.Clear()
            dtGeneral.Clear()

            SBO_Application.StatusBar.SetText("Prepare to load data...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            If LoadPrintData() Then
                SBO_Application.StatusBar.SetText("Data loaded successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                bIsShared = IsSharedFileExist()

                frm = New Hydac_FormViewer
                frm.Name = "PO Details By Vendor Report"
                frm.Text = "PO Details By Vendor Report"
                frm.Report = ReportName.PO_Detail_Proj
                frm.IsShared = bIsShared
                frm.ReportName = ReportName.PO_Detail_Proj
                frm.SharedReportName = g_sReportFilename
                frm.Dataset = ds
                frm.ShowDialog()
                SBO_Application.StatusBar.SetText("Operation ended successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[LoadViewer]:" & ex.Message)
        End Try
    End Sub
    Private Function ValidateParameter() As Boolean
        Try

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.PO_Detail_Proj)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[PO Detail Project].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function LoadPrintData() As Boolean
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim GeneralRow As DataRow
        Dim sWhereQuery As String = ""
        Dim sOrderBy As String = " "
        Dim sSDate As DateTime
        Dim sEDate As DateTime
        Dim sBPFrom As String = ""
        Dim sBPTo As String = ""

        Try
            sWhereQuery = ""
            sOrderBy = " Order By RDR.CardCode, ISNULL(RDD.Project,''), RDR.DocNum "
            oRecordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'Load For General
            GeneralRow = dtGeneral.NewRow
            GeneralRow("CompanyName") = oCompany.CompanyName

            With oForm.DataSources.UserDataSources
                sQuery = "Select RDR.CardCode, RDR.CardName, RDR.DocDate, RDR.DocCur, RDR.DocEntry, RDR.Series, NND.SeriesName , " & _
                " RDR.DocNum, RDD.LineNum, RDD.ItemCode, RDD.Dscription, RDD.Price, RDD.Quantity, RDD.LineTotal, " & _
                " ISNULL(RDD.Project,'') as Project, RDR.CANCELED, ITB.ItmsGrpCod, ITB.ItmsGrpNam, RDD.Rate, RDD.TotalFrgn " & _
                " FROM OPOR RDR " & _
                " INNER JOIN NNM1 NND ON RDR.Series = NND.Series " & _
                " INNER JOIN POR1 RDD ON RDR.DocEntry = RDD.DocEntry " & _
                " INNER JOIN OITM ITM ON RDD.ItemCode = ITM.ItemCode " & _
                " INNER JOIN OITB ITB ON ITB.ItmsGrpCod = ITM.ItmsGrpCod "

                'Get Parameter Value
                Dim date_info As DateTimeFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat()
                Dim FormatDate As String = ""
                Dim sTemp As String = ""
                Dim oEdit As SAPbouiCOM.EditText
                FormatDate = "yyyyMMdd"

                ' ---
                oEdit = oForm.Items.Item("tbDatFr").Specific
                sTemp = oEdit.Value.Trim
                If (sTemp.Length > 0) Then
                    sSDate = DateTime.ParseExact(sTemp, FormatDate, Nothing)
                Else
                    sSDate = DateTime.ParseExact("19900101", FormatDate, Nothing)
                End If

                oEdit = oForm.Items.Item("tbDatTo").Specific
                sTemp = oEdit.Value.Trim
                If (sTemp.Length > 0) Then
                    sEDate = DateTime.ParseExact(sTemp, FormatDate, Nothing)
                Else
                    sEDate = DateTime.Today
                End If

                GeneralRow("DocDateFrom") = sSDate
                GeneralRow("DocDateTo") = sEDate
                ' ---

                Dim sFrDate As String = ""
                Dim sToDate As String = ""
                oEdit = oForm.Items.Item("tbDatFr").Specific
                sFrDate = oEdit.Value.Trim
                oEdit = oForm.Items.Item("tbDatTo").Specific
                sToDate = oEdit.Value.Trim

                If sFrDate.Length > 0 Then
                    sWhereQuery = " WHERE RDR.DocDate >= '" & sFrDate & "' "
                    If sToDate.Length > 0 Then
                        sWhereQuery &= " AND RDR.DocDate <= '" & sToDate & "' "
                    End If
                Else
                    sWhereQuery = " WHERE 1=1 "
                    If sToDate.Length > 0 Then
                        sWhereQuery &= " AND RDR.DocDate <= '" & sToDate & "' "
                    End If
                End If

                ' ----- PROJECT ------------------------------------------------------------
                Dim sProjFr As String = ""
                Dim sProjTo As String = ""

                oEdit = oForm.Items.Item("tbPrjFr").Specific
                sProjFr = oEdit.Value.Trim

                oEdit = oForm.Items.Item("tbPrjTo").Specific
                sProjTo = oEdit.Value.Trim

                GeneralRow("ProjectFrom") = sProjFr
                GeneralRow("ProjectTo") = sProjTo

                If sProjFr.Length > 0 Then
                    sWhereQuery &= " AND RDD.Project >= '" & sProjFr & "' "
                    If sProjTo.Length > 0 Then
                        sWhereQuery &= " AND RDD.Project <= '" & sProjTo & "' "
                    End If
                Else
                    If sProjTo.Length > 0 Then
                        sWhereQuery &= " AND RDD.Project <= '" & sProjTo & "' "
                    End If
                End If

                ' ----- ITEM GROUP ------------------------------------------------------------
                Dim sGroupFr As String = ""
                Dim sGroupTo As String = ""

                oEdit = oForm.Items.Item("tbIGrpFr").Specific
                sGroupFr = oEdit.Value.Trim

                oEdit = oForm.Items.Item("tbIGrpTo").Specific
                sGroupTo = oEdit.Value.Trim

                GeneralRow("ItemGroupFrom") = sGroupFr
                GeneralRow("ItemGroupTo") = sGroupTo

                If sGroupFr.Length > 0 Then
                    sWhereQuery &= " AND ITB.ItmsGrpCod >= '" & sGroupFr & "' "
                    If sGroupTo.Length > 0 Then
                        sWhereQuery &= " AND ITB.ItmsGrpCod <= '" & sGroupTo & "' "
                    End If
                Else
                    If sGroupTo.Length > 0 Then
                        sWhereQuery &= " AND ITB.ItmsGrpCod <= '" & sGroupTo & "' "
                    End If
                End If

                ' ----- VENDOR ------------------------------------------------------------
                oEdit = oForm.Items.Item("tbBPFr").Specific
                sBPFrom = oEdit.Value.Trim

                oEdit = oForm.Items.Item("tbBPTo").Specific
                sBPTo = oEdit.Value.Trim

                GeneralRow("BPFrom") = sBPFrom
                GeneralRow("BPTo") = sBPTo

                If sBPFrom.Length > 0 Then
                    sWhereQuery &= " AND RDR.CardCode >= '" & sBPFrom & "' "
                    If sBPTo.Length > 0 Then
                        sWhereQuery &= " AND RDR.CardCode <= '" & sBPTo & "' "
                    End If
                Else
                    If sBPTo.Length > 0 Then
                        sWhereQuery &= " AND RDR.CardCode <= '" & sBPTo & "' "
                    End If
                End If
                ' -------------------------------------------------------------------------

                If .Item("ckDoc").ValueEx <> "Y" Then
                    sWhereQuery &= " AND RDR.Canceled = 'N' "
                    GeneralRow("IsCancelled") = "No"
                Else
                    GeneralRow("IsCancelled") = "Yes"
                End If

                Dim sExecute As String = ""
                sExecute = sQuery & sWhereQuery & sOrderBy
                Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRec.DoQuery(sExecute)
                If oRec.RecordCount > 0 Then
                    dtGeneral.Rows.Add(GeneralRow)
                    da = New SqlDataAdapter(sExecute, SQLDbConnection)
                    da.SelectCommand.CommandTimeout = 1200
                    da.Fill(dtRpt)
                    da.Dispose()
                    Return True
                Else
                    SBO_Application.StatusBar.SetText("No records found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Return False
                End If
            End With
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[LoadPrintData]:" & ex.Message)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                If pVal.ItemUID = "btPrint" Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                            Return ValidateParameter()
                        End If
                    End If
                End If
            Else
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "btPrint" Then
                            If (oForm.Items.Item(pVal.ItemUID).Enabled) Then
                                Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
                                myThread.SetApartmentState(Threading.ApartmentState.STA)
                                myThread.Start()
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Select Case pVal.ItemUID
                            Case "tbBPFr", "tbBPTo"
                                Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvent.SelectedObjects
                                If (Not oDataTable Is Nothing) Then
                                    Dim sTemp As String = String.Empty
                                    Select Case oCFLEvent.ChooseFromListUID
                                        Case "CFL_BPFr"
                                            sTemp = oDataTable.GetValue("CardCode", 0)
                                            oForm.DataSources.UserDataSources.Item("tbBPFr").ValueEx = sTemp
                                            Exit Select
                                        Case "CFL_BPTo"
                                            sTemp = oDataTable.GetValue("CardCode", 0)
                                            oForm.DataSources.UserDataSources.Item("tbBPTo").ValueEx = sTemp
                                            Exit Select
                                        Case Else
                                            Exit Select
                                    End Select
                                    Return True
                                End If

                            Case "tbPrjFr", "tbPrjTo"
                                Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvent.SelectedObjects
                                If (Not oDataTable Is Nothing) Then
                                    Dim sTemp As String = String.Empty
                                    Select Case oCFLEvent.ChooseFromListUID
                                        Case "CFL_PrjFr"
                                            sTemp = oDataTable.GetValue("PrjCode", 0)
                                            oForm.DataSources.UserDataSources.Item("tbPrjFr").ValueEx = sTemp
                                            Exit Select
                                        Case "CFL_PrjTo"
                                            sTemp = oDataTable.GetValue("PrjCode", 0)
                                            oForm.DataSources.UserDataSources.Item("tbPrjTo").ValueEx = sTemp
                                            Exit Select
                                        Case Else
                                            Exit Select
                                    End Select
                                    Return True
                                End If

                            Case "tbIGrpFr", "tbIGrpTo"
                                Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                                oCFLEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable
                                oDataTable = oCFLEvent.SelectedObjects
                                If (Not oDataTable Is Nothing) Then
                                    Dim sTemp As String = String.Empty
                                    Select Case oCFLEvent.ChooseFromListUID
                                        Case "CFL_ItemGroupFr"
                                            sTemp = oDataTable.GetValue("ItmsGrpCod", 0)
                                            oForm.DataSources.UserDataSources.Item("tbIGrpFr").ValueEx = sTemp
                                            Exit Select
                                        Case "CFL_ItemGroupTo"
                                            sTemp = oDataTable.GetValue("ItmsGrpCod", 0)
                                            oForm.DataSources.UserDataSources.Item("tbIGrpTo").ValueEx = sTemp
                                            Exit Select
                                        Case Else
                                            Exit Select
                                    End Select
                                    Return True
                                End If

                            Case Else
                                'do nothing
                        End Select
                End Select
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ItemEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
