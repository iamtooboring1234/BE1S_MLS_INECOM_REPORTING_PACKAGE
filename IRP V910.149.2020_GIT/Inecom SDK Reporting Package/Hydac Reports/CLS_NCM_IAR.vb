Option Strict Off
Option Explicit On 

Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Threading
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class CLS_NCM_IAR

#Region "Global Variables"
    Private oFormIAR As SAPbouiCOM.Form
    Private sqlConn As SqlConnection
    Private sqlComm As SqlCommand
    Private da As SqlDataAdapter
    Private dr As SqlDataReader
    Private ds As DataSet
    Private dtNCM_OITM As DataTable
    Private StructureFilename As String = ""
    Private ReportFileName As String = ""
    Private g_sLocalCurr As String = ""
    Private sErrMsg As String
    Private lErrCode As Integer
    Private g_sReportFileName As String = ""

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
            Dim oCheck As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            Dim sSQLQuery As String = String.Empty
            Dim sQuery As String = String.Empty
            Dim dTotalValue As Decimal = 0.0
            Dim iValue_A As Integer = 0
            Dim iValue_B As Integer = 0
            Dim iValue_C As Integer = 0
            Dim dValue_A As Decimal = 0.0
            Dim dValue_B As Decimal = 0.0
            Dim dValue_C As Decimal = 0.0

            sQuery = "SELECT U_Query FROM [@NCM_QUERY] WHERE U_Type ='NCM_HYD_OTTL'"
            oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                sSQLQuery = oRecord.Fields.Item(0).Value

                oCheck.DoQuery(sSQLQuery)
                If oCheck.RecordCount > 0 Then
                    oCheck.MoveFirst()
                    dTotalValue = oCheck.Fields.Item(0).Value
                End If
            End If

            oCheck = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oCheck.DoQuery("SELECT * FROM [@NCM_ITEMCAT]")
            If oCheck.RecordCount > 0 Then
                oCheck.MoveFirst()
                While Not oCheck.EoF
                    Select Case oCheck.Fields.Item("Code").Value
                        Case "A"
                            iValue_A = oCheck.Fields.Item("Name").Value
                        Case "B"
                            iValue_B = oCheck.Fields.Item("Name").Value
                        Case "C"
                            iValue_C = oCheck.Fields.Item("Name").Value
                    End Select
                    oCheck.MoveNext()
                End While
            Else
                SBO_Application.StatusBar.SetText("[GenerateDataset] : Please define ABC % in UDT [@NCM_ITEMCAT].", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            dValue_A = dTotalValue * (iValue_A / 100)
            dValue_B = dTotalValue * ((iValue_A + iValue_B) / 100)
            dValue_C = dTotalValue

            ds = New Dataset_Hydac
            dtNCM_OITM = ds.Tables("NCM_OITM")
            If SetSqlConnection() Then
                sQuery = "SELECT U_Query FROM [@NCM_QUERY] WHERE U_Type ='NCM_HYD_OITM'"
                oRecord = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)
                If oRecord.RecordCount > 0 Then
                    oRecord.MoveFirst()
                    sQuery = oRecord.Fields.Item(0).Value
                    sQuery = sQuery.Replace("<<TOTALVALUE>>", dTotalValue)
                    sQuery = sQuery.Replace("<<VALUE_A>>", dValue_A)
                    sQuery = sQuery.Replace("<<VALUE_B>>", dValue_B)
                    sQuery = sQuery.Replace("<<VALUE_C>>", dValue_C)

                    da = New SqlDataAdapter(sQuery, sqlConn)
                    da.Fill(dtNCM_OITM)
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
                    .Text = "Items ABC Analysis Report"
                    .DBUsername = DBUsername
                    .DBPassword = DBPassword
                    .ReportPath = ReportFileName
                    .ReportName = ReportName.Items_ABC_Analysis
                    .IsShared = IsSharedFileExist()
                    .SharedReportName = g_sReportFileName
                    .ReportType = 1
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
            g_sReportFileName = GetSharedFilePath(ReportName.Items_ABC_Analysis)
            If g_sReportFileName <> "" Then
                If IsSharedFilePathExists(g_sReportFileName) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            g_sReportFileName = " "
            SBO_Application.StatusBar.SetText("[IAR].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

End Class
