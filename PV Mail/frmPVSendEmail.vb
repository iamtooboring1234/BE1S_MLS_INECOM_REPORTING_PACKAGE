Imports System.IO

Public Class frmPVSendEmail
#Region "Global Variables"
    Private sQuery As String = ""
    Private BPCode As String = ""

    Private AsAtDate As DateTime
    Private FromDate As DateTime

    Private IsBBF As String
    Private IsGAT As String
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False
    Private cReportName As ReportName

    Private dt As System.Data.DataTable = Nothing
    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
    Dim oMrx As SAPbouiCOM.Matrix
    Dim oCol As SAPbouiCOM.Column

    Private g_s_Selection As String = "0"
#End Region

    Friend Property StatementAsAtDate() As DateTime
        Get
            Return AsAtDate
        End Get
        Set(ByVal value As DateTime)
            AsAtDate = value
        End Set
    End Property
    Friend Property StatementDataTable() As System.Data.DataTable
        Get
            Return dt
        End Get
        Set(ByVal value As System.Data.DataTable)
            dt = value
        End Set
    End Property
    Public Property ReportName() As ReportName
        Get
            Return cReportName
        End Get
        Set(ByVal Value As ReportName)
            cReportName = Value
        End Set
    End Property

#Region "Intialize Application"
    Public Sub New()
        Try

        Catch ex As Exception
            MsgBox("[frmSendEmail].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmPV_SendEmail.srf") Then
            oForm = SBO_Application.Forms.Item("ncmPV_SendEmail")
            AddDataSource()
            SetDatasource()
            InitializeForm()
            oForm.Visible = True
            g_s_Selection = "0"
        Else
            Try
                oForm = SBO_Application.Forms.Item("ncmPV_SendEmail")
                If oForm.Visible = False Then
                    oForm.Close()
                Else
                    InitializeForm()
                    oForm.Select()
                End If
            Catch ex As Exception
                MessageBox.Show("[frmSendEmail].[LoadForm] - " & ex.Message)
            End Try
        End If
    End Sub
    Private Sub AddDataSource()
        With oForm.DataSources.UserDataSources
            .Add("txtAsAt", SAPbouiCOM.BoDataType.dt_DATE, 254)
            .Add("Col00", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 254)
            .Add("Col01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
            .Add("Col02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
            .Add("Col03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            .Add("Col04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5)
            .Add("Col05", SAPbouiCOM.BoDataType.dt_SUM, 254)
            .Add("Col06", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
            .Add("Col07", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 254)
            .Add("Col08", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
            .Add("Col09", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10)
            .Add("Col10", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
        End With
    End Sub
    Private Sub SetDatasource()
        oEdit = oForm.Items.Item("txtAsAt").Specific
        oEdit.DataBind.SetBound(True, String.Empty, "txtAsAt")

        oMrx = oForm.Items.Item("mrxList").Specific

        oCol = oMrx.Columns.Item("Col01")
        oCol.DataBind.SetBound(True, "", "Col01")
        oCol.ValOn = "1"
        oCol.ValOff = "0"
        oCol = oMrx.Columns.Item("Col00")
        oCol.DataBind.SetBound(True, "", "Col00")
        oCol = oMrx.Columns.Item("Col02")
        oCol.DataBind.SetBound(True, "", "Col02")
        oCol = oMrx.Columns.Item("Col03")
        oCol.DataBind.SetBound(True, "", "Col03")
        oCol = oMrx.Columns.Item("Col04")
        oCol.DataBind.SetBound(True, "", "Col04")
        oCol = oMrx.Columns.Item("Col05")
        oCol.DataBind.SetBound(True, "", "Col05")
        oCol = oMrx.Columns.Item("Col06")
        oCol.DataBind.SetBound(True, "", "Col06")
        oCol = oMrx.Columns.Item("Col07")
        oCol.DataBind.SetBound(True, "", "Col07")
        oCol = oMrx.Columns.Item("Col08")
        oCol.DataBind.SetBound(True, "", "Col08")
        oCol = oMrx.Columns.Item("Col09")
        oCol.DataBind.SetBound(True, "", "Col09")
        oCol = oMrx.Columns.Item("Col10")
        oCol.DataBind.SetBound(True, "", "Col10")
    End Sub
    Private Sub InitializeForm()
        Try
            oForm = SBO_Application.Forms.Item("ncmPV_SendEmail")
            oForm.Freeze(True)
        Catch ex As Exception
            Throw ex
        End Try
        oForm.DataSources.UserDataSources.Item("txtAsAt").ValueEx = AsAtDate.ToString("yyyyMMdd")

        Try
            oMrx = oForm.Items.Item("mrxList").Specific
            oMrx.Clear()
            For i As Integer = 1 To dt.Rows.Count Step 1
                With oForm.DataSources.UserDataSources
                    .Item("Col00").ValueEx = i
                    .Item("Col01").ValueEx = dt.Rows(i - 1)("IsEmail")
                    .Item("Col02").ValueEx = dt.Rows(i - 1)("CardCode")
                    .Item("Col03").ValueEx = dt.Rows(i - 1)("CardName")
                    .Item("Col04").ValueEx = dt.Rows(i - 1)("Currency")
                    .Item("Col05").ValueEx = dt.Rows(i - 1)("Balance")
                    .Item("Col06").ValueEx = dt.Rows(i - 1)("EmailTo")
                    .Item("Col07").ValueEx = dt.Rows(i - 1)("Attachment")
                    .Item("Col08").ValueEx = "Ready"
                    .Item("Col09").ValueEx = dt.Rows(i - 1)("DocEntry")
                    .Item("Col10").ValueEx = dt.Rows(i - 1)("DocNum")
                End With
                oMrx.AddRow(1, -1)
            Next
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Logic Function"
    Private Function Save() As Boolean
        Dim sOutput As String = ""
        Try
            Try
                Try
                    oForm = SBO_Application.Forms.Item("ncmPV_SendEmail")
                    oMrx = oForm.Items.Item("mrxList").Specific
                    oForm.Freeze(True)
                Catch ex As Exception
                    Throw ex
                End Try
                Dim s As New clsEmail
                s.GetSetting("PV")
                Dim al As New System.Collections.ArrayList

                For i As Integer = 1 To oMrx.VisualRowCount Step 1
                    sOutput = ""
                    oMrx.GetLineData(i)
                    With oForm.DataSources.UserDataSources
                        If .Item("Col01").ValueEx = "1" Then
                            If Not al.Contains(.Item("Col02").ValueEx) Then
                                al.Add(.Item("Col02").ValueEx)

                                s.Attachment = .Item("Col07").ValueEx
                                s.EmailTo = .Item("Col06").ValueEx
                                s.CardName = .Item("Col03").ValueEx
                                s.DocNum = .Item("Col10").ValueEx

                                If s.SendPVEmail(sOutput, AsAtDate) Then
                                    .Item("Col08").ValueEx = "Sent"
                                Else
                                    .Item("Col08").ValueEx = sOutput
                                End If
                            Else
                                .Item("Col08").ValueEx = "Sent"
                            End If
                        Else
                            .Item("Col08").ValueEx = "Skipped"
                        End If
                    End With
                    oMrx.SetLineData(i)
                Next

                s = Nothing
                oForm.Freeze(False)
                Return True
            Catch ex As Exception
                Throw ex
            End Try
            Return True
        Catch ex As Exception
            Throw New Exception("[frmSendEmail].[Save]" & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
#End Region

#Region "Events Handler"
    Friend Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                        If String.Compare(pVal.ItemUID, "mrxList", True) = 0 Then
                            If String.Compare(pVal.ColUID, "Col07", True) = 0 Then
                                oForm = SBO_Application.Forms.Item("ncmPV_SendEmail")
                                oMrx = oForm.Items.Item("mrxList").Specific
                                BubbleEvent = False
                                oMrx.GetLineData(pVal.Row)
                                Dim sPath As String = oForm.DataSources.UserDataSources.Item("Col07").ValueEx.Trim
                                sPath = sPath

                                Select Case SBO_Application.ClientType
                                    Case SAPbouiCOM.BoClientType.ct_Desktop
                                        System.Diagnostics.Process.Start(sPath)
                                    Case SAPbouiCOM.BoClientType.ct_Browser
                                        SBO_Application.SendFileToBrowser(sPath)
                                End Select
                                Return False
                            End If
                        End If

                    Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                        If pVal.ItemUID = "mrxList" AndAlso pVal.Row = 0 Then
                            If pVal.ColUID = "Col01" Then
                                BubbleEvent = False
                                oForm = SBO_Application.Forms.Item("ncmPV_SendEmail")
                                oMrx = oForm.Items.Item("mrxList").Specific
                                oForm.Freeze(True)
                                Try
                                    Select Case g_s_Selection
                                        Case "0"
                                            g_s_Selection = "1"
                                        Case "1"
                                            g_s_Selection = "0"
                                    End Select

                                    For i As Integer = 1 To oMrx.VisualRowCount Step 1
                                        oMrx.GetLineData(i)
                                        oForm.DataSources.UserDataSources.Item(pVal.ColUID).ValueEx = g_s_Selection
                                        oMrx.SetLineData(i)
                                    Next

                                Catch ex As Exception
                                    Throw ex
                                Finally
                                    oForm.Freeze(False)
                                End Try
                                Return False
                            End If
                        End If
                End Select
            Else 'After Action
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        If String.Compare(pVal.ItemUID, "btnSave", True) = 0 Then
                            BubbleEvent = Save()
                        End If
                End Select 'End Select pval.EventType
            End If 'End If pVal.Before_Action

        Catch ex As Exception
            MsgBox("[frmSendEmail].[ItemEvent]" & vbNewLine & ex.Message)
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class
