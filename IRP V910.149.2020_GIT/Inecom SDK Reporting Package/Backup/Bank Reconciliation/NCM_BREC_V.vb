'' © Copyright © 2007-2019, Inecom Pte Ltd, All rights reserved.
'' =============================================================

Imports SAPbobsCOM
Imports System.Globalization
Imports System.Threading
Imports System.IO
Imports System.Xml
Imports System.Data.Common

Public Class NCM_BREC_V

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oItem As SAPbouiCOM.Item
    Private oStatic As SAPbouiCOM.StaticText
    Private g_sReportFilename As String = String.Empty
    Private dtCommand As System.Data.DataTable
    Private MyDataTable As System.Data.DataTable

    Private ds As System.Data.DataSet
    Public HANADbConnection As DbConnection
    Public _DbProviderFactoryObject As DbProviderFactory
    Public ProviderName As String = "System.Data.Odbc"

    Private g_sFirstAcct As String = ""
    Private g_sLastAcct As String = ""
    Private g_sFirstDate As String = ""
    Private g_sLastDate As String = ""

    Dim g_bIsShared As Boolean = False
    Dim oCheck As SAPbouiCOM.CheckBox
#End Region

#Region "Initialisation"
    Public Sub LoadForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_BREC.srf") Then
            oForm = SBO_Application.Forms.Item("NCM_BREC")
            AddDataSource()
            If (Not oForm.Visible) Then
                oForm.Visible = True
            End If
        Else
            Try
                oForm = SBO_Application.Forms.Item("NCM_BREC")
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
        
            oCFLs = oForm.ChooseFromLists

            oCFLCreationParams = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_GLAccounts
            oCFLCreationParams.UniqueID = "CFL_Acct"
            oCFL = oCFLs.Add(oCFLCreationParams)

            'oCons = New SAPbouiCOM.Conditions
            'oCon = oCons.Add()
            'oCon.Alias = "Postable"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AddChooseFromList] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddDataSource()
        Try
            Dim oEdit As SAPbouiCOM.EditText

            AddChooseFromList()
            With oForm.DataSources.UserDataSources
                .Add("tbDate", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("tbAcct", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("tbFormat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)    'FORMATCODE
                .Add("tbName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
            End With

            oEdit = DirectCast(oForm.Items.Item("tbDate").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, "", "tbDate")
            oEdit = DirectCast(oForm.Items.Item("tbFormat").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, "", "tbFormat")
            oEdit = DirectCast(oForm.Items.Item("tbName").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, "", "tbName")

            oEdit = DirectCast(oForm.Items.Item("tbAcct").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, "", "tbAcct")
            oEdit.ChooseFromListUID = "CFL_Acct"
            oEdit.ChooseFromListAlias = "AcctCode"

            oForm.Items.Item("tbDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AddDataSource] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Print Report"
    Private Function ShowFormatCode(ByVal sAcctCode As String) As String
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim oCheck As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""

            sQuery = " SELECT TOP 1 IFNULL(""EnbSgmnAct"", 'N') FROM " & oCompany.CompanyDB & ".""CINF"" "
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                If oRecord.Fields.Item(0).Value = "N" Then
                    Return sAcctCode
                Else
                    sQuery = "  SELECT IFNULL(""Segment_0"", '') || IFNULL('-' || ""Segment_1"", '') || IFNULL('-' || ""Segment_2"", '') || IFNULL('-' || ""Segment_3"", '') || IFNULL('-' || ""Segment_4"", '') || IFNULL('-' || ""Segment_5"", '') || IFNULL('-' || ""Segment_6"", '') || IFNULL('-' || ""Segment_7"", '') || IFNULL('-' || ""Segment_8"", '') || IFNULL('-' || ""Segment_9"", '') "
                    sQuery &= " FROM " & oCompany.CompanyDB & ".""OACT"" "
                    sQuery &= " WHERE ""AcctCode"" = '" & sAcctCode & "'"

                    oCheck.DoQuery(sQuery)
                    If oCheck.RecordCount > 0 Then
                        oCheck.MoveFirst()
                        Return oCheck.Fields.Item(0).Value
                    End If
                End If
            End If

            Return sAcctCode
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ShowFormatCode] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.BankReconciliation)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[GL Listing].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        Try
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\BREC_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
            End If
            oRec = Nothing

            ' ===============================================================================
            ' get the folder of the report of the current DB Name
            ' set to local
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\BREC\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\BREC_" & sCurrDate & "_" & sCurrTime & ".pdf"
            ' ===============================================================================

            If GenerateDataset() Then
                Dim frm As New Hydac_FormViewer
                frm.ReportName = ReportName.BankReconciliation
                frm.Text = "Bank Reconciliation Report"
                frm.Name = "Bank Reconciliation Report"
                frm.IsShared = IsSharedFileExist()
                frm.SharedReportName = g_sReportFilename
                frm.BankAccount = oForm.DataSources.UserDataSources.Item("tbAcct").ValueEx
                frm.BankDate = oForm.DataSources.UserDataSources.Item("tbDate").ValueEx
                frm.Dataset = ds
                frm.ExportPath = sFinalFileName
                Select Case SBO_Application.ClientType
                    Case SAPbouiCOM.BoClientType.ct_Desktop
                        frm.ClientType = "D"
                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.ClientType = "S"
                End Select

                Select Case SBO_Application.ClientType
                    Case SAPbouiCOM.BoClientType.ct_Desktop
                        frm.ShowDialog()

                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.OPEN_HANADS_BANKRECONCILIATION()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select

            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Function GenerateDataset() As Boolean
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sRecord As String = ""
            Dim sQuery As String = ""
            Dim sBankAccount As String = ""
            Dim sBankDate As String = ""

            sBankAccount = oForm.DataSources.UserDataSources.Item("tbAcct").ValueEx
            sBankDate = oForm.DataSources.UserDataSources.Item("tbDate").ValueEx

            ds = New XML_GLL
            dtCommand = ds.Tables("NCM_BANKRECON;1")

            sQuery = "CALL """ & oCompany.CompanyDB & """.""NCM_BANKRECON"" ('" & sBankAccount & "','" & sBankDate & "') "
            'sQuery = "CALL ""ECS_ENT_LIVE_30062017"".""NCM_BANKRECON"" ('_SYS00000000359','20151130') "

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                If SetSqlConnection() Then
                    Using HANADbConnection
                        If HANADbConnection.State <> ConnectionState.Open Then
                            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
                            HANADbConnection = _DbProviderFactoryObject.CreateConnection()
                            HANADbConnection.ConnectionString = connStr
                            HANADbConnection.Open()
                        End If

                        Dim da As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
                        Dim command As DbCommand = HANADbConnection.CreateCommand()
                        _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)

                        MyDataTable = New DataTable
                        command.CommandText = sQuery
                        command.ExecuteNonQuery()
                        da.SelectCommand = command
                        da.Fill(dtCommand)
                    End Using

                    Return True
                Else
                    SBO_Application.StatusBar.SetText("[GenerateDataset1] : Failed to open SQL connection.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            Else
                SBO_Application.StatusBar.SetText("[GenerateDataset2] : No records found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
 
            oRecord = Nothing
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GenerateDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function SetSqlConnection() As Boolean
        Try
            connStr = "DRIVER={HDBODBC32};UID=" & global_DBUsername & ";PWD=" & global_DBPassword & ";SERVERNODE=" & oCompany.Server & ";DATABASE=" & oCompany.CompanyDB & ""

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            HANADbConnection = _DbProviderFactoryObject.CreateConnection()
            HANADbConnection.ConnectionString = connStr
            HANADbConnection.Open()

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[SetSqlConnection] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function PrintReport() As Boolean
        Try
            Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
            myThread.SetApartmentState(ApartmentState.STA)
            myThread.Start()
            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[Print]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function ValidateInput() As Boolean
        Try
            Dim sInputDate As String = ""
            Dim sInputAcct As String = ""

            sInputDate = oForm.DataSources.UserDataSources.Item("tbDate").ValueEx.ToString.Trim
            sInputAcct = oForm.DataSources.UserDataSources.Item("tbAcct").ValueEx.ToString.Trim

            If sInputDate = "" Then
                SBO_Application.StatusBar.SetText("[ValidateInput]: Input 'As At Date' is blank.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If sInputAcct = "" Then
                SBO_Application.StatusBar.SetText("[ValidateInput]: Input 'Account Code' is blank.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[ValidateInput]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                    Case "btPrint"
                        If pval.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If ValidateInput() Then
                                BubbleEvent = False
                                SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                PrintReport()
                            End If
                         
                        End If
                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Select Case pval.ItemUID
                            Case "tbAcct"
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim sCFL_ID As String = ""
                                Dim sCode As String = ""
                                Dim sName As String = ""
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim oDataTable As SAPbouiCOM.DataTable

                                oCFLEvento = pval
                                sCFL_ID = oCFLEvento.ChooseFromListUID
                                oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                                oDataTable = oCFLEvento.SelectedObjects

                                Try
                                    sCode = oDataTable.GetValue("AcctCode", 0)
                                    sName = oDataTable.GetValue("AcctName", 0)
                                Catch ex As Exception

                                End Try
                             
                                oForm.DataSources.UserDataSources.Item("tbAcct").ValueEx = sCode
                                oForm.DataSources.UserDataSources.Item("tbName").ValueEx = sName
                                oForm.DataSources.UserDataSources.Item("tbFormat").ValueEx = ShowFormatCode(sCode)

                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            SBO_Application.StatusBar.SetText("[ItemEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
