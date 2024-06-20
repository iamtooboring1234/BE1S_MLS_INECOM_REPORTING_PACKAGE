Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class GST_FrmViewer
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
        Me.GST_CrtViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'GST_CrtViewer
        '
        Me.GST_CrtViewer.ActiveViewIndex = -1
        Me.GST_CrtViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GST_CrtViewer.Location = New System.Drawing.Point(0, 0)
        Me.GST_CrtViewer.Name = "GST_CrtViewer"
        Me.GST_CrtViewer.ReportSource = Nothing
        Me.GST_CrtViewer.Size = New System.Drawing.Size(292, 273)
        Me.GST_CrtViewer.TabIndex = 0
        '
        'GST_FrmViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.GST_CrtViewer)
        Me.Name = "GST_FrmViewer"
        Me.Text = "GST Report"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Private Variable"
    Dim cReportName As ReportName = ReportName.GST
    Dim bIsReportExternal As Boolean = False
    Dim sSGSTCode As String = String.Empty
    Dim sEGSTCode As String = String.Empty
    Dim sSDate As Date = DateTime.Now
    Dim sEDate As Date = DateTime.Now
    Dim sReportType As String = String.Empty
    Dim sCurrencyType As String = String.Empty
    Dim sShowGL As String = String.Empty
    Dim g_sSharedReportName As String = String.Empty
    Dim cDBPassword As String = SubMain.DBPassword
    Dim cDBUser As String = SubMain.DBUsername
    Dim sCrystalDateFormat As String = "DateValue({0},{1},{2})"
    Dim ds As DataSet
    Private g_sGST_Curr As String = ""
    Private g_sCompanyName As String = ""
    Private cInputTax As String = "I"
    Private oExportType As CrystalDecisions.Shared.ExportFormatType = ExportFormatType.PortableDocFormat
    Private g_bIsExport As Boolean = False
    Private g_sExportPath As String = ""
    Private g_sClientType As String = "D"

#End Region

#Region "Properties"
    Friend Property InputTax() As String
        Get
            Return cInputTax
        End Get
        Set(ByVal value As String)
            cInputTax = value
        End Set
    End Property
    Friend Property GST_Currency() As String
        Get
            Return g_sGST_Curr
        End Get
        Set(ByVal value As String)
            g_sGST_Curr = value
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
    Public Property ClientType() As String
        Get
            Return g_sClientType
        End Get
        Set(ByVal Value As String)
            g_sClientType = Value
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
    Public Property IsExportReport() As Boolean
        Get
            Return g_bIsExport
        End Get
        Set(ByVal Value As Boolean)
            g_bIsExport = Value
        End Set
    End Property
    Public Property ExportPath() As String
        Get
            Return g_sExportPath
        End Get
        Set(ByVal Value As String)
            g_sExportPath = Value
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
    Public Property DBUsernameViewer() As String
        Get
            Return cDBUser
        End Get
        Set(ByVal Value As String)
            cDBUser = Value
        End Set
    End Property
    Public Property DBPasswordViewer() As String
        Get
            Return cDBPassword
        End Get
        Set(ByVal Value As String)
            cDBPassword = Value
        End Set
    End Property
    Public Property StartGST() As String
        Get
            Return sSGSTCode
        End Get
        Set(ByVal Value As String)
            sSGSTCode = Value
        End Set
    End Property
    Public Property EndGST() As String
        Get
            Return sEGSTCode
        End Get
        Set(ByVal Value As String)
            sEGSTCode = Value
        End Set
    End Property
    Public Property StartDate() As Date
        Get
            Return sSDate
        End Get
        Set(ByVal Value As Date)
            sSDate = Value
        End Set
    End Property
    Public Property EndDate() As Date
        Get
            Return sEDate
        End Get
        Set(ByVal Value As Date)
            sEDate = Value
        End Set
    End Property
    Public Property ReportType() As String
        Get
            Return sReportType
        End Get
        Set(ByVal Value As String)
            sReportType = Value
        End Set
    End Property
    Public Property CurrencyType() As String
        Get
            Return sCurrencyType
        End Get
        Set(ByVal Value As String)
            sCurrencyType = Value
        End Set
    End Property

    Public Property ShowGLAccount() As String
        Get
            Return sShowGL
        End Get
        Set(ByVal Value As String)
            sShowGL = Value
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

    Friend Sub OpenGSTReport()
        Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
        Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
        Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim rpt As CrystalDecisions.CrystalReports.Engine.ReportDocument

        crConnectionInfo.ServerName = oCompany.Server
        crConnectionInfo.DatabaseName = oCompany.CompanyDB
        crConnectionInfo.UserID = DBUsername
        crConnectionInfo.Password = DBPassword

        If bIsReportExternal Then
            rpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New RPT_GST_GLDET
        End If

        With rpt
            .DataDefinition.FormulaFields.Item("StartGST").Text = SetStringIntoCrystalFormula(sSGSTCode)
            .DataDefinition.FormulaFields.Item("EndGST").Text = SetStringIntoCrystalFormula(sEGSTCode)
            .DataDefinition.FormulaFields.Item("StartDate").Text = String.Format(sCrystalDateFormat, sSDate.Year.ToString("000#"), sSDate.Month.ToString("0#"), sSDate.Day.ToString("0#"))
            .DataDefinition.FormulaFields.Item("EndDate").Text = String.Format(sCrystalDateFormat, sEDate.Year.ToString("000#"), sEDate.Month.ToString("0#"), sEDate.Day.ToString("0#"))
            .DataDefinition.FormulaFields.Item("ReportType").Text = sReportType
            .DataDefinition.FormulaFields.Item("ShowGLAccount").Text = SetStringIntoCrystalFormula(sShowGL)
            .DataDefinition.FormulaFields.Item("ShowCurrency").Text = sCurrencyType
            .DataDefinition.FormulaFields.Item("GSTCurrency").Text = "'" & g_sGST_Curr & "'"
            .DataDefinition.FormulaFields.Item("CompanyName").Text = "'" & oCompany.CompanyName & "'"

            Select Case cInputTax
                Case "I"
                    .DataDefinition.FormulaFields.Item("InputTax").Text = "'Input Tax'"
                Case "O"
                    .DataDefinition.FormulaFields.Item("InputTax").Text = "'Output Tax'"
                Case "A"
                    .DataDefinition.FormulaFields.Item("InputTax").Text = "'Input & Output Tax'"
            End Select

            .SetDataSource(ds)
            .Refresh()
        End With

        If g_sClientType = "S" Then
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
    Private Sub GST_CrtViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Select Case cReportName
                Case ReportName.GST
                    OpenGSTReport()
            End Select
        Catch ex As Exception
            MessageBox.Show("[GST_FrmViewer].[frmViewer_Load]" & vbNewLine & ex.Message)
            Me.Close()
        End Try
    End Sub
End Class
