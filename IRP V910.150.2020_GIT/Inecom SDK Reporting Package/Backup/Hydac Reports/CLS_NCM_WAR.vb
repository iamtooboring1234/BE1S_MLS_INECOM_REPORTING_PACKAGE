Option Strict Off
Option Explicit On

Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class CLS_NCM_WAR

#Region "Global Variables"
    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oFormWAR As SAPbouiCOM.Form

    Private sqlConn As New SqlConnection
    Private da As SqlDataAdapter
    Private ds As DataSet
    Private dtNCM_WAR As DataTable

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

            'Quantity Transacted = Invoice – Credit Note + Prod Order Issue – Prod Order Receipt + Stock Adj Issue – Stock Adj Receipt
            'If Quantity falls in between LAST 12 - 18 months,  Weighted Value = Quantity (x 1)
            'If Quantity falls in between LAST 6 – 12 months,   Weighted Value = Quantity (x 2)
            'If Quantity falls in between LAST 0 – 6 months,    Weighted Value = Quantity (x 3)
            'Quantity Demand per week = round down (Total of 3 Weight Value) / 156 to the nearest integer
            'As for MRP Forecast table structure, I suggest you ask Harianto directly.

            ds = New Dataset_Hydac
            dtNCM_WAR = ds.Tables("NCM_WAR")
            If SetSqlConnection() Then
                sQuery = "SELECT U_Query FROM [@NCM_QUERY] WHERE U_Type ='NCM_HYD_WAR'"
                oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)
                If oRecord.RecordCount > 0 Then
                    oRecord.MoveFirst()
                    sQuery = oRecord.Fields.Item(0).Value

                    da = New SqlDataAdapter(sQuery, sqlConn)
                    da.Fill(dtNCM_WAR)
                Else
                    SBO_Application.StatusBar.SetText("[GenerateDataset] : Running Query NCM_HYD_WAR, the result is blank. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    .Server = oCompany.Server
                    .Database = oCompany.CompanyDB
                    .Text = "Weighted Average Demand Report"
                    .DBUsername = DBUsername
                    .DBPassword = DBPassword
                    .ReportPath = ReportFileName
                    .ReportName = ReportName.Weighted_Average_Demand
                    .IsShared = IsSharedFileExist()
                    .SharedReportName = g_sReportFileName
                    .ReportType = 3
                    .Dataset = ds
                    .ShowDialog()
                End With
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[LoadViewer] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFileName = GetSharedFilePath(ReportName.Weighted_Average_Demand)
            If g_sReportFileName <> "" Then
                If IsSharedFilePathExists(g_sReportFileName) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFileName = " "
            SBO_Application.StatusBar.SetText("[WAR].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
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

#End Region

#Region "Event Handler"
    Public Function ItemEvent(ByRef pval As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pval.BeforeAction = False Then
                '' DO NOTHING
            Else
                Select Case pval.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If pval.ItemUID = "btPrint" Then
                            ' -----------------------------------------------------------------------------------------
                            BubbleEvent = False
                            SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            PrintReport()
                            ' -----------------------------------------------------------------------------------------
                        End If
                End Select
            End If
        Catch ex As Exception
            oFormWAR.Freeze(False)
            SBO_Application.StatusBar.SetText("[ItemEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
