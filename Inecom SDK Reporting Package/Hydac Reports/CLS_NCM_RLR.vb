Option Strict Off
Option Explicit On

Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class CLS_NCM_RLR

#Region "Global Variables"
    Private oFormRLR As SAPbouiCOM.Form
    Private sqlConn As SqlConnection
    Private sqlComm As SqlCommand
    Private da As SqlDataAdapter
    Private dr As SqlDataReader
    Private ds As DataSet

    Private dtNCM_OITM As DataTable
    Private dtNCM_DOC As DataTable

    Private StructureFilename As String = ""
    Private ReportFileName As String = ""
    Private g_sLocalCurr As String = ""
    Private g_sReportFileName As String = ""

    Private sErrMsg As String
    Private lErrCode As Integer
#End Region

#Region "Constructors"
    Public Sub New()
        MyBase.new()
    End Sub
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
#End Region

#Region "General Functions"
    Friend Sub LoadForm()
        Dim oEdit As SAPbouiCOM.EditText

        Try
            If LoadFromXML("Inecom_SDK_Reporting_Package." & FRM_NCM_RLR & ".srf") Then ' Loading .srf file
                SBO_Application.StatusBar.SetText("Loading form...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oFormRLR = SBO_Application.Forms.Item(FRM_NCM_RLR)
                oFormRLR.SupportedModes = -1

                oFormRLR.EnableMenu(MenuID.Remove, False)
                oFormRLR.EnableMenu(MenuID.Find, False)
                oFormRLR.EnableMenu(MenuID.Add, False)
                oFormRLR.EnableMenu(MenuID.Delete, False)
                oFormRLR.EnableMenu(MenuID.Paste, True)
                oFormRLR.EnableMenu(MenuID.Copy, True)
                oFormRLR.EnableMenu(MenuID.Cut, True)
                oFormRLR.EnableMenu(MenuID.Undo, False)

                With oFormRLR.DataSources.UserDataSources
                    .Add("uNoweek", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
                    .Add("uSmooth", SAPbouiCOM.BoDataType.dt_PRICE)
                End With

                oEdit = oFormRLR.Items.Item("tbNoweek").Specific
                oEdit.DataBind.SetBound(True, "", "uNoweek")

                oEdit = oFormRLR.Items.Item("tbSmooth").Specific
                oEdit.DataBind.SetBound(True, "", "uSmooth")
                ' -----------------------------------------------------
                oFormRLR.DataSources.UserDataSources.Item("uNoweek").ValueEx = 0
                oFormRLR.DataSources.UserDataSources.Item("uSmooth").ValueEx = 0

                oFormRLR.Visible = True
                SBO_Application.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Else
                ' Loading .srf file failed most likely it is because the form is already opened
                Try
                    oFormRLR = SBO_Application.Forms.Item(FRM_NCM_RLR)
                    If oFormRLR.Visible Then
                        oFormRLR.Select()
                    Else
                        oFormRLR.Close()
                    End If
                Catch ex As Exception
                End Try
            End If

        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadForm] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Report"
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
    Private Function GenerateDataset() As Boolean
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            Dim myRow As DataRow
            Dim myRows() As DataRow
            Dim myLoopRow As DataRow
            Dim myLoopRows() As DataRow

            Dim dSmoothQty As Decimal = 0
            Dim dCurrValue As Decimal = 0
            Dim dTotalMov As Decimal = 0
            Dim dTotalQty As Decimal = 0
            Dim dStep1 As Decimal = 0.0
            Dim iStep1 As Integer = 0
            Dim dFactor As Decimal = oFormRLR.DataSources.UserDataSources.Item("uSmooth").ValueEx
            Dim iNoweek As Integer = oFormRLR.DataSources.UserDataSources.Item("uNoWeek").ValueEx

            Dim iA_value As Integer = 0
            Dim iB_value As Integer = 0
            Dim iC_value As Integer = 0

            ds = New Dataset_Hydac
            dtNCM_OITM = ds.Tables("NCM_OITM")
            dtNCM_DOC = ds.Tables("NCM_DOC")

            If SetSqlConnection() Then
                ' -------------------------------------------------------------------------

                sQuery = "SELECT U_Query FROM [@NCM_QUERY] WHERE U_Type ='NCM_HYD_OITM_RLR'"
                oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)
                If oRecord.RecordCount > 0 Then
                    oRecord.MoveFirst()
                    sQuery = oRecord.Fields.Item(0).Value

                    da = New SqlDataAdapter(sQuery, sqlConn)
                    da.Fill(dtNCM_OITM)

                    If Not (dtNCM_OITM Is Nothing) Then
                        myRows = dtNCM_OITM.Select()

                        For Each myRow In myRows
                            myRow.BeginEdit()
                            dTotalMov = 0.0
                            dTotalQty = 0.0
                            dStep1 = 0.0
                            iStep1 = 0

                            dTotalMov = Convert.ToDecimal(myRow("IGESum")) + Convert.ToDecimal(myRow("DLNSum")) + Convert.ToDecimal(myRow("RDNSum")) + Convert.ToDecimal(myRow("RINSum")) + Convert.ToDecimal(myRow("INVSum"))
                            dTotalQty = Convert.ToDecimal(myRow("IGEQty")) + Convert.ToDecimal(myRow("DLNQty")) + Convert.ToDecimal(myRow("RDNQty")) + Convert.ToDecimal(myRow("RINQty")) + Convert.ToDecimal(myRow("INVQty"))

                            'Step 1: Round down of (Last 12 mths Sales Qty x Smoothing Factor / Total # of Movement) to nearest integer.
                            If dTotalMov > 0 Then
                                dStep1 = (dTotalQty * dFactor) / dTotalMov
                                iStep1 = Convert.ToInt32(dStep1)

                                sQuery = "SELECT U_Query FROM [@NCM_QUERY] WHERE U_Type ='NCM_HYD_DOC'"
                                oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                                oRecord.DoQuery(sQuery)
                                If oRecord.RecordCount > 0 Then
                                    oRecord.MoveFirst()
                                    sQuery = oRecord.Fields.Item(0).Value
                                    sQuery = sQuery.Replace("<<ITEMCODE>>", myRow("ItemCode"))

                                    da = New SqlDataAdapter(sQuery, sqlConn)
                                    da.Fill(dtNCM_DOC)
                                End If

                                'Step2: Loop through all documents involved in obtaining Total # of Movement (Invoice,Credit Note,Prod Order Issue,Prod Order Rcpt,Stock Issue,Stock rcpt)
                                '       For each Document, compare Item Qty (A) against Step 1 value
                                '       If A > Step 1 value, B = (Step 1 Value – A)
                                '       If A < Step 1 value, B = A
                                '       Repeat for all documents
                                'Step3: Add all B’s value to Total Movement Qty to give C. (C = Sum of all B)
                                'Step4: Divide C by Total # of Movements = New Smoothed Average Movement per Item.

                                iC_value = 0
                                If Not (dtNCM_DOC Is Nothing) Then
                                    myLoopRows = dtNCM_DOC.Select()

                                    For Each myLoopRow In myLoopRows
                                        myLoopRow.BeginEdit()
                                        iA_value = 0
                                        iB_value = 0

                                        iA_value = Convert.ToInt32(myLoopRow("Quantity"))
                                        Select Case iA_value
                                            Case Is > iStep1
                                                iB_value = iA_value - iStep1
                                            Case Is <= iStep1
                                                iB_value = iA_value
                                        End Select

                                        iC_value += iB_value
                                        myLoopRow.BeginEdit()
                                    Next
                                End If

                                myRow("SmoothQty") = iC_value / dTotalMov

                            Else
                                myRow("SmoothQty") = 0
                            End If

                            myRow.EndEdit()
                        Next
                        dtNCM_OITM.AcceptChanges()
                    End If
                    '' -------------------------------------------------------------------------
                Else
                    SBO_Application.StatusBar.SetText("[GenerateDataset] : No records found.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oRecord = Nothing
                    Return False
                End If

                sqlConn.Close()
                oRecord = Nothing
            Else
                SBO_Application.StatusBar.SetText("[GenerateDataset] : Failed to open SQL connection.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oRecord = Nothing
                Return False
            End If

            oRecord = Nothing
            Return True
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GenerateDataset] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Private Sub LoadViewer()
        Try
            Dim ofrmViewer As New Hydac_FormViewer
            If GenerateDataset() Then
                With ofrmViewer
                    .SmoothFactor = oFormRLR.DataSources.UserDataSources.Item("uSmooth").ValueEx
                    .NoOfWeek = oFormRLR.DataSources.UserDataSources.Item("uNoweek").ValueEx
                    .Text = "ReOrder Level Recommendation Report"
                    .Server = oCompany.Server
                    .Database = oCompany.CompanyDB
                    .DBUsername = DBUsername
                    .DBPassword = DBPassword
                    .ReportPath = ReportFileName
                    .ReportName = ReportName.ReOrder_Level_Recommendation
                    .IsShared = IsSharedFileExist()
                    .SharedReportName = g_sReportFileName
                    .ReportType = 2
                    .Dataset = ds
                    .ShowDialog()
                End With
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Friend Function PrintReport() As Boolean
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
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFileName = GetSharedFilePath(ReportName.ReOrder_Level_Recommendation)
            If g_sReportFileName <> "" Then
                If IsSharedFilePathExists(g_sReportFileName) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFileName = " "
            SBO_Application.StatusBar.SetText("[RLR].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = False Then
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "btPrint" Then
                            If oFormRLR.DataSources.UserDataSources.Item("uNoweek").ValueEx <= 0 Then
                                SBO_Application.StatusBar.SetText("Invalid No. of Weeks. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                ' -----------------------------------------------------------------------------------------
                                SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                PrintReport()
                                ' -----------------------------------------------------------------------------------------
                            End If
                        End If
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
