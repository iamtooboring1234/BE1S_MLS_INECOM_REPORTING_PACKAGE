Imports SAPbobsCOM
Imports System.Globalization
Imports System.Threading
Imports System.IO
Imports System.Xml
Imports System.Data.Common

Public Class NCM_GLL_V

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oItem As SAPbouiCOM.Item
    Private oStatic As SAPbouiCOM.StaticText
    Private g_sReportFilename As String = String.Empty
    Private dtCommand As System.Data.DataTable
    Private ds As System.Data.DataSet
    'Private sqlConn As SqlConnection
    'Private sqlComm As SqlCommand
    'Private da As SqlDataAdapter
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
        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_GLL.srf") Then
            oForm = SBO_Application.Forms.Item("NCM_GLL")
            AddDataSource()
            If (Not oForm.Visible) Then
                oForm.Visible = True
            End If
        Else
            Try
                oForm = SBO_Application.Forms.Item("NCM_GLL")
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
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_GLAccounts
            oCFLCreationParams.UniqueID = "CFL_AcctFr"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            'oCon.Alias = "Levels"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "5"
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            'oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)


            oCFLCreationParams = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = SAPbouiCOM.BoLinkedObject.lf_GLAccounts
            oCFLCreationParams.UniqueID = "CFL_AcctTo"
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = New SAPbouiCOM.Conditions
            oCon = oCons.Add()
            'oCon.Alias = "Levels"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "5"
            'oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            'oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)



        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AddChooseFromList] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub AddDataSource()
        Try
            Dim oEdit As SAPbouiCOM.EditText
            Dim oCbox As SAPbouiCOM.ComboBox

            AddChooseFromList()
            With oForm.DataSources.UserDataSources
                .Add("tbDateFr", SAPbouiCOM.BoDataType.dt_DATE, 254)
                .Add("tbDateTo", SAPbouiCOM.BoDataType.dt_DATE, 1)
                .Add("tbAcctFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("tbAcctTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("tbAcctNmFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbAcctNmTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbAcctFrFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbAcctFrTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("cbBal", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            End With

            oEdit = DirectCast(oForm.Items.Item("tbDateFr").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbDateFr")
            oEdit = DirectCast(oForm.Items.Item("tbDateTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbDateTo")
            oEdit = DirectCast(oForm.Items.Item("tbAcctFr").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbAcctFr")
            oEdit.ChooseFromListUID = "CFL_AcctFr"
            oEdit.ChooseFromListAlias = "AcctCode"
            oEdit = DirectCast(oForm.Items.Item("tbAcctTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbAcctTo")
            oEdit.ChooseFromListUID = "CFL_AcctTo"
            oEdit.ChooseFromListAlias = "AcctCode"
            oEdit = DirectCast(oForm.Items.Item("tbAcctNmFr").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbAcctNmFr")
            oEdit = DirectCast(oForm.Items.Item("tbAcctNmTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbAcctNmTo")
            oEdit = DirectCast(oForm.Items.Item("tbAcctFrFr").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbAcctFrFr")
            oEdit = DirectCast(oForm.Items.Item("tbAcctFrTo").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "tbAcctFrTo")

            oCbox = oForm.Items.Item("cbBal").Specific
            oCbox.DataBind.SetBound(True, "", "cbBal")
            oCbox.ValidValues.Add("N", "NO")
            oCbox.ValidValues.Add("Y", "YES")
            oCbox.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.DataSources.UserDataSources.Item("cbBal").ValueEx = "N"

            oForm.Items.Item("tbDateFr").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

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
    Private Function GenerateDataset() As Boolean
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sRecord As String = ""
            Dim sQuery As String = ""
            Dim sOpBalance As String = "N"
            sOpBalance = oForm.DataSources.UserDataSources.Item("cbBal").ValueEx
            ' ------------------------------------------------------------------------------

            g_sFirstAcct = ""
            g_sLastAcct = ""

            ds = New XML_GLL
            dtCommand = ds.Tables("Command")

            sQuery = " SELECT TOP 1 ""AcctCode"" FROM " & oCompany.CompanyDB & ".""OACT"" ORDER BY ""AcctCode"" ASC"
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                g_sFirstAcct = oRecord.Fields.Item(0).Value
            End If

            sQuery = " SELECT TOP 1 ""AcctCode"" FROM " & oCompany.CompanyDB & ".""OACT"" ORDER BY ""AcctCode"" DESC  "
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                g_sLastAcct = oRecord.Fields.Item(0).Value
            End If

            If oForm.DataSources.UserDataSources.Item("tbAcctFr").ValueEx <> "" Then
                g_sFirstAcct = oForm.DataSources.UserDataSources.Item("tbAcctFr").ValueEx
            End If

            If oForm.DataSources.UserDataSources.Item("tbAcctTo").ValueEx <> "" Then
                g_sLastAcct = oForm.DataSources.UserDataSources.Item("tbAcctTo").ValueEx
            End If

            g_sFirstDate = "19900101"
            ' ------------------------------------------------------------------------------
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(" SELECT IFNULL(CAST(YEAR(NOW()) AS varchar) || Right('00'|| CAST(MONTH(NOW()) AS varchar(2)), 2) || Right('00'|| CAST(DAYOFMONTH(NOW()) AS varchar(2)), 2), '') AS ""CurrDate"" FROM DUMMY ")
            g_sLastDate = oRecord.Fields.Item(0).Value

            If oForm.DataSources.UserDataSources.Item("tbDateFr").ValueEx <> "" Then
                g_sFirstDate = oForm.DataSources.UserDataSources.Item("tbDateFr").ValueEx
            End If

            If oForm.DataSources.UserDataSources.Item("tbDateTo").ValueEx <> "" Then
                g_sLastDate = oForm.DataSources.UserDataSources.Item("tbDateTo").ValueEx
            End If

            sQuery = "CALL " & oCompany.CompanyDB & ".""NCM_GL_LISTING"" ('" & g_sFirstAcct.Trim & "','" & g_sLastAcct.Trim & "','" & g_sFirstDate.Trim & "','" & g_sLastDate.Trim & "','" & sOpBalance & "')"
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
            SBO_Application.StatusBar.SetText("[GenerateDatasetX] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.GL_Listing)
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
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\GL_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
            End If
            oRec = Nothing

             ' ===============================================================================
            ' get the folder of AR SOA of the current DB Name
            ' set to local
            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\GL\" & oCompany.CompanyDB
            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
            If Not di.Exists Then
                di.Create()
            End If
            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
            sFinalFileName = di.FullName & "\GL_" & sCurrDate & "_" & sCurrTime & ".pdf"
            ' ===============================================================================

            If GenerateDataset() Then
                Dim frm As New VIEWER_GLL
                frm.Text = "GL Listing Report"
                frm.Name = "GL Listing Report"
                frm.IsReportExternal = IsSharedFileExist()
                frm.SharedReportName = g_sReportFilename
                frm.AccountFr = oForm.DataSources.UserDataSources.Item("tbAcctFrFr").ValueEx
                frm.AccountTo = oForm.DataSources.UserDataSources.Item("tbAcctFrTo").ValueEx
                frm.StartDate = g_sFirstDate
                frm.EndDate = g_sLastDate
                frm.OpBalance = oForm.DataSources.UserDataSources.Item("cbBal").ValueEx
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
                        frm.OpenGLLReport()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select

            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function SetSqlConnection() As Boolean
        Try
           
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
#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = True Then
                Select Case pval.ItemUID
                    Case "btPrint"
                        If pval.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            BubbleEvent = False
                            SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            PrintReport()
                        End If
                End Select
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Select Case pval.ItemUID
                            Case "tbAcctFr", "tbAcctTo"
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
                                Select Case pval.ItemUID
                                    Case "tbAcctFr"
                                        oForm.DataSources.UserDataSources.Item("tbAcctFr").ValueEx = sCode
                                        oForm.DataSources.UserDataSources.Item("tbAcctNmFr").ValueEx = sName
                                        oForm.DataSources.UserDataSources.Item("tbAcctFrFr").ValueEx = ShowFormatCode(sCode)
                                    Case "tbAcctTo"
                                        oForm.DataSources.UserDataSources.Item("tbAcctTo").ValueEx = sCode
                                        oForm.DataSources.UserDataSources.Item("tbAcctNmTo").ValueEx = sName
                                        oForm.DataSources.UserDataSources.Item("tbAcctFrTo").ValueEx = ShowFormatCode(sCode)
                                End Select
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
