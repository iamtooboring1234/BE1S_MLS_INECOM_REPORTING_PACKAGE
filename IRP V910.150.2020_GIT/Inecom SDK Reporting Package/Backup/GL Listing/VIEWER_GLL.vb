Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class VIEWER_GLL
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GST_CrtViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(VIEWER_GLL))
        Me.GST_CrtViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'GST_CrtViewer
        '
        Me.GST_CrtViewer.AccessibleDescription = Nothing
        Me.GST_CrtViewer.AccessibleName = Nothing
        Me.GST_CrtViewer.ActiveViewIndex = -1
        resources.ApplyResources(Me.GST_CrtViewer, "GST_CrtViewer")
        Me.GST_CrtViewer.BackgroundImage = Nothing
        Me.GST_CrtViewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.GST_CrtViewer.Font = Nothing
        Me.GST_CrtViewer.Name = "GST_CrtViewer"
        Me.GST_CrtViewer.SelectionFormula = ""
        Me.GST_CrtViewer.ViewTimeSelectionFormula = ""
        '
        'VIEWER_GLL
        '
        Me.AccessibleDescription = Nothing
        Me.AccessibleName = Nothing
        resources.ApplyResources(Me, "$this")
        Me.BackgroundImage = Nothing
        Me.Controls.Add(Me.GST_CrtViewer)
        Me.Font = Nothing
        Me.Icon = Nothing
        Me.Name = "VIEWER_GLL"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Private Variable"
    Dim cReportName As ReportName = ReportName.GST
    Dim bIsReportExternal As Boolean = False
    Dim sSGSTCode As String = String.Empty
    Dim sEGSTCode As String = String.Empty

    Dim sReportType As String = String.Empty
    Dim sCurrencyType As String = String.Empty
    Dim sShowGL As String = String.Empty
    Dim g_sSharedReportName As String = String.Empty
    Dim cDBPassword As String = SubMain.DBPassword
    Dim cDBUser As String = SubMain.DBUsername
    Dim sCrystalDateFormat As String = "DateValue({0},{1},{2})"
    Dim ds As DataSet
    Dim cOpBalance As String = "N"

    Private g_sParamAccount As String = ""
    Private g_sParamAccountTo As String = ""
    Private g_sParamDate As String = ""
    Private g_sParamDateTo As String = ""
    Private g_sExportPath As String = ""
    Private cClientType As String = "D"

#End Region

#Region "Properties"
    Friend Property ExportPath() As String
        Get
            Return g_sExportPath
        End Get
        Set(ByVal Value As String)
            g_sExportPath = Value
        End Set
    End Property
    Friend Property ClientType() As String
        Get
            Return cClientType
        End Get
        Set(ByVal Value As String)
            cClientType = Value
        End Set
    End Property
    Public Property OpBalance() As String
        Get
            Return cOpBalance
        End Get
        Set(ByVal Value As String)
            cOpBalance = Value
        End Set
    End Property
    Public Property IsReportExternal() As Boolean
        Get
            Return bIsReportExternal
        End Get
        Set(ByVal Value As Boolean)
            bIsReportExternal = Value
        End Set
    End Property
    Public Property SharedReportName() As String
        Get
            Return g_sSharedReportName
        End Get
        Set(ByVal Value As String)
            g_sSharedReportName = Value
        End Set
    End Property
    Public Property AccountFr() As String
        Get
            Return g_sParamAccount
        End Get
        Set(ByVal Value As String)
            g_sParamAccount = Value
        End Set
    End Property
    Public Property AccountTo() As String
        Get
            Return g_sParamAccountTo
        End Get
        Set(ByVal Value As String)
            g_sParamAccountTo = Value
        End Set
    End Property
    Public Property StartDate() As String
        Get
            Return g_sParamDate
        End Get
        Set(ByVal Value As String)
            g_sParamDate = Value
        End Set
    End Property
    Public Property EndDate() As String
        Get
            Return g_sParamDateTo
        End Get
        Set(ByVal Value As String)
            g_sParamDateTo = Value
        End Set
    End Property
    Public Property Dataset() As System.Data.DataSet
        Get
            Return ds
        End Get
        Set(ByVal Value As System.Data.DataSet)
            ds = Value
        End Set
    End Property
#End Region

    Friend Sub OpenGLLReport()

        Dim rpt As CrystalDecisions.CrystalReports.Engine.ReportDocument
        If bIsReportExternal Then
            rpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New RPT_GLL
        End If

        With rpt
            .DataDefinition.FormulaFields.Item("OpBalance").Text = "'" & cOpBalance & "'"
            .DataDefinition.FormulaFields.Item("ParamAccount").Text = SetStringIntoCrystalFormula(g_sParamAccount)
            .DataDefinition.FormulaFields.Item("ParamAccountTo").Text = SetStringIntoCrystalFormula(g_sParamAccountTo)
            .DataDefinition.FormulaFields.Item("ParamDate").Text = "'" & g_sParamDate.Substring(6, 2) & "." & g_sParamDate.Substring(4, 2) & "." & g_sParamDate.Substring(0, 4) & "'"
            .DataDefinition.FormulaFields.Item("ParamDateTo").Text = "'" & g_sParamDateTo.Substring(6, 2) & "." & g_sParamDateTo.Substring(4, 2) & "." & g_sParamDateTo.Substring(0, 4) & "'"
            .SetDataSource(ds)
            .Refresh()
        End With

        If cClientType = "S" Then
            ' Web Browser
            ' generate pdf, put to a standard local folder, with specific name.
            ' call the pdf out.

            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = g_sExportPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
        Else
            'Desktop
            GST_CrtViewer.ReportSource = rpt
            GST_CrtViewer.Show()
        End If

    End Sub
    Private Function SetStringIntoCrystalFormula(ByVal InputString As String) As String
        If (InputString Is Nothing) Then
            Return String.Empty
        Else
            Return "'" & InputString.Replace("'", "''") & "'"
        End If
    End Function
    Private Sub VIEWER_GLL_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GST_CrtViewer.Load
        Try
            OpenGLLReport()
        Catch ex As Exception
            MessageBox.Show("[GLLViewer].[frmViewer_Load]" & vbNewLine & ex.Message)
            Me.Close()
        End Try
    End Sub
End Class
