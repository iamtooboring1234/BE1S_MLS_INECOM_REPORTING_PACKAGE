Imports System.IO
Imports System.Data.Common

Public Class NCM_GST

#Region "Global Variables"
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oItem As SAPbouiCOM.Item
    Private oStatic As SAPbouiCOM.StaticText
    Private oPictureBox As SAPbouiCOM.PictureBox
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Private oCheck As SAPbouiCOM.CheckBox
    Private dsRpt As System.Data.DataSet
    Private dt_OVTG As System.Data.DataTable
    Private g_sGST_Curr As String = ""
    Private Timer1 As New System.Windows.Forms.Timer
    Private g_iSecond As Integer = 0
    Private g_sGSTRunningDate As String = ""

#End Region

#Region "Initialization"
    Public Sub New()
        AddHandler Timer1.Tick, AddressOf Timer_Tick
    End Sub
    Private Sub Timer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        g_iSecond += 1
        oStatic.Caption = "Processing " & g_iSecond & " seconds ..."
    End Sub

    Public Sub ShowForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.SRF_NCM_RPT_GST.srf") Then
            oForm = SBO_Application.Forms.Item("NCM_RPT_GSTT")
            AddDataSource()
            If (Not oForm.Visible) Then
                oForm.Visible = True
            End If
            SetupChooseFromList()
        Else
            Try
                oForm = SBO_Application.Forms.Item("NCM_RPT_GSTT")
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
        Try
            With oForm.DataSources.UserDataSources
                .Add("txtSGST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtEGST", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 8)
                .Add("txtSDate", SAPbouiCOM.BoDataType.dt_DATE, 254)
                .Add("txtEDate", SAPbouiCOM.BoDataType.dt_DATE, 1)
                .Add("cboType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("cboCurr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("cboGL", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("cbInput", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            End With

            oEdit = DirectCast(oForm.Items.Item("txtSGST").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtSGST")
            oEdit = DirectCast(oForm.Items.Item("txtEGST").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtEGST")
            oEdit = DirectCast(oForm.Items.Item("txtSDate").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtSDate")

            oForm.DataSources.UserDataSources.Item("txtSDate").ValueEx = DateTime.Now.ToString("yyyy0101")

            oEdit = DirectCast(oForm.Items.Item("txtEDate").Specific, SAPbouiCOM.EditText)
            oEdit.DataBind.SetBound(True, String.Empty, "txtEDate")
            Dim sTemp As String = String.Empty
            sTemp = DateTime.Now.Year.ToString("000#") + DateTime.Now.Month.ToString("0#") + DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString("0#")

            oForm.DataSources.UserDataSources.Item("txtEDate").ValueEx = sTemp

            oCombo = DirectCast(oForm.Items.Item("cbInput").Specific, SAPbouiCOM.ComboBox)
            oCombo.DataBind.SetBound(True, String.Empty, "cbInput")
            oCombo.ValidValues.Add("A", "All")
            oCombo.ValidValues.Add("I", "Input Tax")
            oCombo.ValidValues.Add("O", "Output Tax")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.DataSources.UserDataSources.Item("cbInput").ValueEx = "A"

            oCombo = DirectCast(oForm.Items.Item("cboType").Specific, SAPbouiCOM.ComboBox)
            oCombo.DataBind.SetBound(True, String.Empty, "cboType")
            oCombo.ValidValues.Add("0", "Summary")
            oCombo.ValidValues.Add("1", "Details")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.DataSources.UserDataSources.Item("cboType").ValueEx = "0"

            oCombo = DirectCast(oForm.Items.Item("cboCurr").Specific, SAPbouiCOM.ComboBox)
            oCombo.DataBind.SetBound(True, String.Empty, "cboCurr")
            oCombo.ValidValues.Add("0", "LC & FC")
            oCombo.ValidValues.Add("1", "LC & SC")
            oCombo.ValidValues.Add("2", "SC & FC")
            oCombo.ValidValues.Add("3", "GST Curr & LC")
            oCombo.ValidValues.Add("4", "GST Curr & SC")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.DataSources.UserDataSources.Item("cboCurr").ValueEx = "0"

            oCombo = DirectCast(oForm.Items.Item("cboGL").Specific, SAPbouiCOM.ComboBox)
            oCombo.DataBind.SetBound(True, String.Empty, "cboGL")
            oCombo.ValidValues.Add("0", "Yes")
            oCombo.ValidValues.Add("1", "No")
            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            oForm.DataSources.UserDataSources.Item("cboGL").ValueEx = "0"

        Catch ex As Exception
            SBO_Application.MessageBox("[NCM_GST].[AddDataSource]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function SetChooseFromListConditions(ByVal myVal As String, ByVal compareVal As String, ByVal cflUID As String) As Boolean
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFL = oForm.ChooseFromLists.Item(cflUID)
        oCons = New SAPbouiCOM.Conditions
        Try
            Select Case cflUID
                Case "cflSGST"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    End If

                    Exit Select
                Case "cflEGST"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "Code"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    End If
                    Exit Select
                Case Else
                    Throw New Exception("Invalid Choose from list. UID#" & cflUID)
                    Exit Select
            End Select
            oCon = oCons.Add()
            oCFL.SetConditions(oCons)
            Return True
        Catch ex As Exception
            Throw New Exception("[NCM_GST].[SetChooseFromList] " & vbNewLine & ex.Message)
            Return False
        End Try
        Return False
    End Function
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oForm.ChooseFromLists

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "5"
            oCFLCreation.UniqueID = "cflSGST"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtSGST").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflSGST"
            oEditLn.ChooseFromListAlias = "Code"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "5"
            oCFLCreation.UniqueID = "cflEGST"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oForm.Items.Item("txtEGST").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflEGST"
            oEditLn.ChooseFromListAlias = "Code"
        Catch ex As Exception
            Throw New Exception("[NCM_GST].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Function GetLocalCurrency() As String
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""MainCurncy"" FROM " & oCompany.CompanyDB & ".""OADM"" ")
            If (oRec.RecordCount > 0) Then
                Return oRec.Fields.Item(0).Value.ToString()
            End If
            oRec = Nothing
            Return ""
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[NCM_GST].[GetLocalCurrency] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End Try
    End Function
    Private Function GetGSTCurrency() As Boolean
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            g_sGST_Curr = ""
            oRec.DoQuery("SELECT IFNULL(""U_GSTCURR"",'') FROM " & oCompany.CompanyDB & ".""@NCM_NEW_SETTING"" ")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                g_sGST_Curr = oRec.Fields.Item(0).Value
                oRec = Nothing
                Return True
            Else
                oRec = Nothing
                Return False
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[NCM_GST].[GetGSTCurrency] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.GST)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[GST].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Private Sub LoadViewer()
        oForm.Items.Item("btnPrint").Enabled = False
        Try
            Dim frm As New GST_FrmViewer
            Dim bIsContinue As Boolean = False
            Dim TempDate As DateTime
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sTempDirectory As String = ""
            Dim sPathFormat As String = "{0}\GST_{1}.pdf"
            Dim sCurrDate As String = ""
            Dim sFinalExportPath As String = ""
            Dim sFinalFileName As String = ""
            Dim sCurrTime As String = DateTime.Now.ToString("HHMMss")

            g_sGSTRunningDate = ""
            oRec.DoQuery("SELECT CAST(current_timestamp AS NVARCHAR(24)), TO_CHAR(current_timestamp, 'YYYYMMDD') FROM DUMMY")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                g_sGSTRunningDate = Convert.ToString(oRec.Fields.Item(0).Value).Trim
                sCurrDate = Convert.ToString(oRec.Fields.Item(1).Value).Trim
            End If

            Try
                ' ===============================================================================
                ' get the folder of AR SOA of the current DB Name
                ' set to local
                sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\GST\" & oCompany.CompanyDB
                Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                If Not di.Exists Then
                    di.Create()
                End If
                sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                sFinalFileName = di.FullName & "\GST_" & sCurrDate & "_" & sCurrTime & ".pdf"
                ' ===============================================================================

                Timer1.Interval = 1000
                Timer1.Enabled = True
                If ExecuteProcedure() Then
                    If PrepareDataSet() Then
                        bIsContinue = True

                        With oForm.DataSources.UserDataSources
                            TempDate = DateTime.ParseExact(.Item("txtSDate").ValueEx, "yyyyMMdd", Nothing)
                            frm.StartDate = TempDate
                            TempDate = DateTime.ParseExact(.Item("txtEDate").ValueEx, "yyyyMMdd", Nothing)
                            frm.EndDate = TempDate

                            frm.StartGST = .Item("txtSGST").ValueEx
                            frm.EndGST = .Item("txtEGST").ValueEx
                            frm.ReportType = .Item("cboType").ValueEx
                            frm.CurrencyType = .Item("cboCurr").ValueEx
                            frm.InputTax = .Item("cbInput").ValueEx
                        End With

                        oCombo = oForm.Items.Item("cboGL").Specific
                        If (Not oCombo.Selected Is Nothing) Then
                            frm.ShowGLAccount = oCombo.Selected.Description
                        Else
                            frm.ShowGLAccount = "Yes"
                        End If

                        frm.IsReportExternal = g_bIsShared
                        frm.SharedReportName = g_sReportFilename
                        frm.GST_Currency = g_sGST_Curr
                        frm.Dataset = dsRpt
                        frm.ReportName = ReportName.GST
                        frm.ExportPath = sFinalFileName
                        Select Case SBO_Application.ClientType
                            Case SAPbouiCOM.BoClientType.ct_Desktop
                                frm.ClientType = "D"
                            Case SAPbouiCOM.BoClientType.ct_Browser
                                frm.ClientType = "S"
                        End Select

                        Timer1.Enabled = False
                    End If
                End If
            Catch ex As Exception
                Throw ex
            Finally
                oForm.Items.Item("btnPrint").Enabled = True
                oStatic.Caption = String.Empty
            End Try
            If bIsContinue Then
                Select Case SBO_Application.ClientType
                    Case SAPbouiCOM.BoClientType.ct_Desktop
                        frm.ShowDialog()

                    Case SAPbouiCOM.BoClientType.ct_Browser
                        frm.OpenGSTReport()

                        If File.Exists(sFinalFileName) Then
                            SBO_Application.SendFileToBrowser(sFinalFileName)
                        End If
                End Select
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            Timer1.Enabled = False
        End Try
    End Sub
    Private Function ExecuteProcedure() As Boolean
        GetGSTCurrency()
        g_bIsShared = IsSharedFileExist()

        Dim sSGST As String = ""
        Dim sEGST As String = ""
        Dim sDate_Fr As String = ""
        Dim sDate_To As String = ""
        Dim sQuery As String = ""
        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sSDate As DateTime
        Dim sEDate As DateTime
        Dim sTemp As String = ""

        Try
            oEdit = oForm.Items.Item("txtSGST").Specific
            sSGST = oEdit.Value
            oEdit = oForm.Items.Item("txtEGST").Specific
            sEGST = oEdit.Value

            sTemp = oForm.DataSources.UserDataSources.Item("txtSDate").ValueEx
            sDate_Fr = sTemp
            If (sTemp.Length > 0) Then
                sSDate = DateTime.ParseExact(sTemp, "yyyyMMdd", Nothing)
            Else
                sSDate = DateTime.ParseExact("19821223", "yyyyMMdd", Nothing)
            End If

            sTemp = oForm.DataSources.UserDataSources.Item("txtEDate").ValueEx
            sDate_To = sTemp
            If (sTemp.Length > 0) Then
                sEDate = DateTime.ParseExact(sTemp, "yyyyMMdd", Nothing)
            Else
                sEDate = DateTime.Today
            End If

            oStatic = oForm.Items.Item("lbStatus").Specific
            oStatic.Caption = "Executing Store Procedure. Please wait..."

            sQuery = "  CALL NCM_RPT_GST ("
            sQuery &= " '" & g_sGSTRunningDate & oCompany.UserName & "',"
            sQuery &= " '" & sSGST & "',"
            sQuery &= " '" & sEGST & "',"
            sQuery &= " '" & sDate_Fr & "',"
            sQuery &= " '" & sDate_To & "',"
            sQuery &= " '" & oCompany.UserName & "')"
            oRec.DoQuery(sQuery)

            oStatic.Caption = ""
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[NCM_GST][ExecuteProcedure]:" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function ValidateParameter() As Boolean
        Try
            oForm.ActiveItem = "txtSGST"
            Dim oRec As SAPbobsCOM.Recordset = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty
            Dim sQuery As String = String.Empty

            With oForm.DataSources.UserDataSources
                sStart = .Item("txtSGST").ValueEx
                sEnd = .Item("txtEGST").ValueEx
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("GST Code From is greater than GST Code To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.ActiveItem = "txtSGST"
                        Return False
                    End If
                End If
                sStart = .Item("txtSDate").ValueEx
                If (sStart.Length = 0) Then
                    SBO_Application.StatusBar.SetText("Please enter start date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtSDate"
                    Return False
                End If
                sEnd = .Item("txtEDate").ValueEx
                If (sEnd.Length = 0) Then
                    SBO_Application.StatusBar.SetText("Please enter end date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oForm.ActiveItem = "txtEDate"
                    Return False
                End If
                If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                    If (String.Compare(sStart, sEnd) > 0) Then
                        SBO_Application.StatusBar.SetText("Date From is greater than Date To", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oForm.ActiveItem = "txtSDate"
                        Return False
                    End If
                End If
            End With
            SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[NCM_GST].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function PrepareDataSet() As Boolean
        Try
            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            HANADbConnection = _DbProviderFactoryObject.CreateConnection()
            HANADbConnection.ConnectionString = connStr
            HANADbConnection.Open()

            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim HANAda As DbDataAdapter = _DbProviderFactoryObject.CreateDataAdapter()
            Dim HANAcmd As DbCommand
            Dim sQueryFLn As String = ""
            Dim sSelect As String = ""

            dsRpt = New GST_Report_XML
            Dim dtRpt As System.Data.DataTable = dsRpt.Tables("DS_RPT_GST")

            ' commented out to resolve Showing GL Account
            ' sQueryFLn &= " ON T1.""VATACCTCODE"" = T2.""Account"" "

            sQueryFLn &= " SELECT T1.* "
            sQueryFLn &= " FROM """ & oCompany.CompanyDB & """.""@NCM_RPT_GSTY"" T1 "
            sQueryFLn &= " LEFT OUTER JOIN """ & oCompany.CompanyDB & """.""OVTG"" T2 ON T1.""VATCODE"" = T2.""Code"" "
            sQueryFLn &= " WHERE  T1.""USERNAME"" = '" & oCompany.UserName & "' "
            sQueryFLn &= " AND  T1.""GSTRUNNINGTIME"" = '" & g_sGSTRunningDate & oCompany.UserName & "' "

            Select Case oForm.DataSources.UserDataSources.Item("cbInput").ValueEx
                Case "I"
                    sQueryFLn &= " AND T2.""Category"" = 'I' "
                Case "O"
                    sQueryFLn &= " AND T2.""Category"" = 'O' "
            End Select

            sSelect = " SELECT ""Code"", ""Name"" FROM """ & oCompany.CompanyDB & """.""OVTG"" "
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sSelect)
            If oRec.RecordCount > 0 Then
                HANAcmd = HANADbConnection.CreateCommand()
                HANAcmd.CommandText = sSelect
                HANAcmd.ExecuteNonQuery()
                HANAda.SelectCommand = HANAcmd
                HANAda.Fill(dsRpt, "DS_OVTG")
            End If

            sSelect = " SELECT ""CompnyName"", IFNULL(""CompnyAddr"",'') ""CompnyAddr"" ,IFNULL(""PrintHeadr"",'') ""PrintHeadr"" , IFNULL(""Phone1"",'') ""Phone1"", IFNULL(""Phone2"",'') ""Phone2"", IFNULL(""Fax"",'') ""Fax"", IFNULL(""E_Mail"",'') ""E_Mail"", ""MainCurncy"", ""SysCurrncy"", IFNULL(""TaxIdNum"",'') ""TaxIdNum"", IFNULL(""RevOffice"",'') ""RevOffice"", IFNULL(""FreeZoneNo"",'') ""FreeZoneNo"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sSelect)
            If oRec.RecordCount > 0 Then
                HANAcmd = HANADbConnection.CreateCommand()
                HANAcmd.CommandText = sSelect
                HANAcmd.ExecuteNonQuery()
                HANAda.SelectCommand = HANAcmd
                HANAda.Fill(dsRpt, "DS_OADM")
            End If

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQueryFLn)
            If (oRec.RecordCount > 0) Then
                HANAcmd = HANADbConnection.CreateCommand()
                HANAcmd.CommandText = sQueryFLn
                HANAcmd.ExecuteNonQuery()
                HANAda.SelectCommand = HANAcmd
                HANAda.Fill(dsRpt, "DS_RPT_GST")
            Else
                HANADbConnection.Close()
                SBO_Application.StatusBar.SetText("No record found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            End If

            HANADbConnection.Close()
            oRec = Nothing
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[NCM_GST][PrepareDataSet]:" & ex.ToString)
            Return False
        End Try
        Return False
    End Function
#End Region

#Region "Events Handler"
    Public Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
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
                            Case "cflSGST"
                                sTemp = oDataTable.GetValue("Code", 0)
                                oForm.DataSources.UserDataSources.Item("txtSGST").ValueEx = sTemp
                                Exit Select
                            Case "cflEGST"
                                sTemp = oDataTable.GetValue("Code", 0)
                                oForm.DataSources.UserDataSources.Item("txtEGST").ValueEx = sTemp
                                Exit Select
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
            SBO_Application.StatusBar.SetText("[NCM_GST].[ItemEvent]" & vbNewLine & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
