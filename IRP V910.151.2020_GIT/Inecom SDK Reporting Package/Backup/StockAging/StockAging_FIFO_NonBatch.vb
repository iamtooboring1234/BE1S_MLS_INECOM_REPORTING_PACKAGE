Option Strict Off
Option Explicit On 

Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class StockAging_FIFO_NonBatch

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

    Private g_sReportFilename As String
    Private g_sXSDFilename As String = String.Empty
    Private g_iSecond As Integer
    Dim aTimer As New System.Windows.Forms.Timer
    Private g_bIsShared As Boolean = False
    Dim statusThread As System.Threading.Thread
    Dim myThread As System.Threading.Thread
    Dim oStaticLn As SAPbouiCOM.StaticText

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
            Dim oCbox As SAPbouiCOM.ComboBox
            Dim oChck As SAPbouiCOM.CheckBox
            Dim oLink As SAPbouiCOM.LinkedButton
            Dim sCode1 As String = String.Empty
            Dim sCode2 As String = String.Empty
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            If LoadFromXML("Inecom_SDK_Reporting_Package." & FRM_NCM_FIFO1 & ".srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading Form...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_FIFO1)
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
                oCbox.ValidValues.Add("1", "Warehouse")
                oCbox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)

                PopulateDate()
                SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
                oFormFIFO1.Visible = True
            Else
                Try
                    oFormFIFO1 = SBO_Application.Forms.Item(FRM_NCM_FIFO1)
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
            oCon.BracketOpenNum = 2
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "EvalSystem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "F"
            oCon.BracketCloseNum = 2
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
            oCon.BracketOpenNum = 2
            oCon.Alias = "InvntItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "EvalSystem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "F"
            oCon.BracketCloseNum = 2
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
            sQuery = " SELECT ISNULL(CAST(YEAR(GetDate()) AS VARCHAR) + Right(Replicate('0',2) + CAST(MONTH(GetDate()) AS VARCHAR(2)),2) + Right(Replicate('0',2) + CAST(DAY(GetDate()) AS VARCHAR(2)),2),'') As CurrDate "
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                sCurrDate = oRecord.Fields.Item(0).Value
                oFormFIFO1.DataSources.UserDataSources.Item("uAsDate").ValueEx = sCurrDate
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[PopulateDate] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Print Report"
    Private Function IsSharedFileExist() As Boolean
        Try
            Select Case oFormFIFO1.DataSources.UserDataSources.Item("cboRptType").ValueEx
                Case "0"
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_FIFO_SUMMARY)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
                Case Else
                    g_sReportFilename = GetSharedFilePath(ReportName.SAR_FIFO_DETAILS)
                    If g_sReportFilename <> "" Then
                        If IsSharedFilePathExists(g_sReportFilename) Then
                            Return True
                        End If
                    End If
            End Select
            Return False
        Catch ex As Exception
            g_sReportFilename = " "
            SBO_Application.StatusBar.SetText("[StockAgeing_FIFO].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Function SetSqlConnection() As Boolean
        Try
            sqlConn = New SqlConnection(DBConnString)
            sqlConn.Open()
            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[SetSqlConnection] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub Timer_Tick(ByVal sender As Object, ByVal e As System.EventArgs)
        g_iSecond += 1
        SBO_Application.StatusBar.SetText("Processing " & g_iSecond & " seconds ...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    End Sub
    Private Function ExecuteProcedure() As Boolean
        Try
            Dim sQuery As String = String.Empty
            Dim SQLCommand As System.Data.SqlClient.SqlCommand
            Dim sItemFr, sItemTo, sWareFr, sWareTo, sAsDate As String
            Dim sItmGrpFr, sItmGrpTo As String
            With oFormFIFO1.DataSources.UserDataSources
                sItemFr = .Item("uItemFr").ValueEx
                sItemTo = .Item("uItemTo").ValueEx
                sWareFr = .Item("uWareFr").ValueEx
                sWareTo = .Item("uWareTo").ValueEx
                sAsDate = .Item("uAsDate").ValueEx
                sItmGrpFr = .Item("tbItmGrpFr").ValueEx
                sItmGrpTo = .Item("tbItmGrpTo").ValueEx
            End With

            sQuery = "EXECUTE NCM_SP_SAR_FIFO1 "
            sQuery &= "'" & sAsDate & "', "
            sQuery &= "'" & oCompany.UserSignature & "', "
            sQuery &= "'" & sItemFr & "', "
            sQuery &= "'" & sItemTo & "', "
            sQuery &= "'" & sWareFr & "', "
            sQuery &= "'" & sWareTo & "',  "
            sQuery &= "'" & sItmGrpFr & "', "
            sQuery &= "'" & sItmGrpTo & "'"

            Try
                statusThread.Start()
                SQLCommand = sqlConn.CreateCommand
                SQLCommand.CommandText = sQuery
                SQLCommand.CommandTimeout = 36000   '2011-Jan-12 (Jessie) Increase for SeiWoo (3600 --> 36000)
                SQLCommand.CommandType = CommandType.Text
                SQLCommand.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                SBO_Application.StatusBar.SetText("SP is not completed successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Throw ex
            End Try
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[ExecProd] : " & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Private Function GenerateRecords() As Boolean
        Try
            Dim oQuery As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            bIsExportToExcel = IIf(oFormFIFO1.DataSources.UserDataSources.Item("chkExcel").ValueEx = "Y", True, False)

            Dim oComboLn As SAPbouiCOM.ComboBox
            oComboLn = oFormFIFO1.Items.Item("cboRptType").Specific

            If (Not oComboLn.Selected Is Nothing) Then
                sRptType = oComboLn.Selected.Value
            Else
                sRptType = "0"
            End If

            ds = New SAR_FIFO1
            dtFIFO = ds.Tables("DS_FIFO")
            dtExportS = ds.Tables("Excel_Output")
            '' ----------------------------------------------------------------------

            sQuery = "SELECT U_Query FROM [@NCM_QUERY] WHERE U_Type ='NCM_SAR_FIFO1'"
            oQuery.DoQuery(sQuery)
            If oQuery.RecordCount > 0 Then
                oQuery.MoveFirst()
                sQuery = oQuery.Fields.Item(0).Value

                sQuery = sQuery.Replace("<<USERSIGN>>", oCompany.UserSignature)
                sQuery = sQuery.Replace("<<RUNDATE>>", "'" & oFormFIFO1.DataSources.UserDataSources.Item("uAsDate").ValueEx & "'")
                sQuery = sQuery.Trim
                da = New SqlDataAdapter(sQuery, sqlConn)
                da.Fill(dtFIFO)


                'sQuery = "SELECT ""ItemCode"", ""ItemName"", ""InvntryUom"" FROM """ & oCompany.CompanyDB & """.""OITM"" "
                'dtOITM = ds.Tables("OCRD")
                'HANAcmd = dbConn.CreateCommand()
                'HANAcmd.CommandText = sQuery
                'HANAcmd.ExecuteNonQuery()
                'HANAda.SelectCommand = HANAcmd
                'HANAda.Fill(dtOITM)
                ''--------------------------------------------------------
                ''OADM (Company Details)
                ''--------------------------------------------------------
                'sQuery = "SELECT ""CompnyAddr"",""CompnyName"",""E_Mail"",""Fax"",""FreeZoneNo"",""RevOffice"",""Phone1"" FROM """ & oCompany.CompanyDB & """.""OADM"" "
                'dtOADM = ds.Tables("OADM")
                'HANAcmd = dbConn.CreateCommand()
                'HANAcmd.CommandText = sQuery
                'HANAcmd.ExecuteNonQuery()
                'HANAda.SelectCommand = HANAcmd
                'HANAda.Fill(dtOADM)

                '' ----------------------------------------------------------------------
                oQuery = Nothing
                Return True
            End If

            SBO_Application.StatusBar.SetText("[GenerateRecords] : Cannot find query in [@NCM_QUERY].", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oQuery = Nothing
            Return False
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
            If (i = 4) Then
                i = 0
            Else
                i = i + 1
            End If
        Loop
    End Sub
    Private Sub PrintSAR_FIFO_NonBatch()
        Dim sFinalExportPath As String = ""
        Dim sFinalFileName As String = ""

        Try
            statusThread = New System.Threading.Thread(AddressOf PrintStatus)
            Try
                oFormFIFO1.Items.Item("btPrint").Enabled = False
                g_iSecond = 0
                g_bIsShared = IsSharedFileExist()

                If SetSqlConnection() Then
                    If ExecuteProcedure() Then
                        If GenerateRecords() Then
                            '' To print Stock Aging Report FIFO Non-Batch Items
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
                            ' get the folder of the current DB Name
                            ' set to local
                            sTempDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\STOCK\" & oCompany.CompanyDB
                            Dim di As New System.IO.DirectoryInfo(sTempDirectory)
                            If Not di.Exists Then
                                di.Create()
                            End If
                            sFinalExportPath = String.Format(sPathFormat, di.FullName, sCurrDate & "_" & sCurrTime)
                            sFinalFileName = di.FullName & "\STOCK_" & sCurrDate & "_" & sCurrTime & ".pdf"
                            ' ===============================================================================

                            Dim oSARVwr As New SARVwr
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
                                .ExcludeZeroBalance = oFormFIFO1.DataSources.UserDataSources.Item("uExcl").ValueEx.ToString
                                .AsAtDate = oFormFIFO1.DataSources.UserDataSources.Item("uAsDate").ValueEx
                                .IsShared = g_bIsShared
                                .IsExcel = bIsExportToExcel
                                .SharedReportName = g_sReportFilename
                                .ItemGroupFrom = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpFr").ValueEx
                                .ItemGroupTo = oFormFIFO1.DataSources.UserDataSources.Item("tbItmGrpTo").ValueEx

                                If (oFormFIFO1.DataSources.UserDataSources.Item("cboRptType").ValueEx = "0") Then
                                    .ReportType = "S"
                                    .ReportName = ReportName.SAR_FIFO_SUMMARY
                                Else
                                    .ReportType = "D"
                                    .ReportName = ReportName.SAR_FIFO_DETAILS
                                End If

                                If (oFormFIFO1.DataSources.UserDataSources.Item("cboGrpBy").ValueEx = "0") Then
                                    .GroupBy = "ItemCode"
                                Else
                                    .GroupBy = "WhsCode"
                                End If

                                If (bIsExportToExcel) Then
                                    If (String.Compare(sRptType, "0", True) = 0) Then
                                        .OpenSummaryReport_FIFO()
                                    Else
                                        .OpenDetailReport_FIFO()
                                    End If
                                Else

                                    Select Case SBO_Application.ClientType
                                        Case SAPbouiCOM.BoClientType.ct_Desktop
                                            oSARVwr.ShowDialog()

                                        Case SAPbouiCOM.BoClientType.ct_Browser
                                            If (String.Compare(sRptType, "0", True) = 0) Then
                                                oSARVwr.OpenSummaryReport_FIFO()
                                            Else
                                                oSARVwr.OpenDetailReport_FIFO()
                                            End If

                                            If File.Exists(sFinalFileName) Then
                                                SBO_Application.SendFileToBrowser(sFinalFileName)
                                            End If
                                    End Select

                                End If
                            End With
                        End If
                    End If
                    sqlConn.Close()
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
            oStaticLn.Caption = String.Empty
        End Try
    End Sub
#End Region

#Region "Event Handlers"
    Friend Function ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
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
