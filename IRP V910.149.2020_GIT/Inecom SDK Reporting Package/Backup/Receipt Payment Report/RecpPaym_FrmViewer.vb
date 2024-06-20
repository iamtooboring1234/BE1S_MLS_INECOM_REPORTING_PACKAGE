'JN Added V03.28.2007 to print Receipt & Payment List (requested by POSITIVE)

Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine

Public Class RecpPaym_FrmViewer
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
    Friend WithEvents RecpPaym_crViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.RecpPaym_crViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'RecpPaym_crViewer
        '
        Me.RecpPaym_crViewer.ActiveViewIndex = -1
        Me.RecpPaym_crViewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.RecpPaym_crViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RecpPaym_crViewer.Location = New System.Drawing.Point(0, 0)
        Me.RecpPaym_crViewer.Name = "RecpPaym_crViewer"
        Me.RecpPaym_crViewer.SelectionFormula = ""
        Me.RecpPaym_crViewer.Size = New System.Drawing.Size(292, 273)
        Me.RecpPaym_crViewer.TabIndex = 0
        Me.RecpPaym_crViewer.ViewTimeSelectionFormula = ""
        '
        'RecpPaym_FrmViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.RecpPaym_crViewer)
        Me.Name = "RecpPaym_FrmViewer"
        Me.Text = "RecpPaym_frmViewer"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Global Variables"
    Private cReport As ReportName
    Private cReportPath As String = String.Empty
    Private cReportDataSet As DataSet
    Private cUserCode As String = ""
    Private cCompanyName As String = SBO_Application.Company.Name
    Private g_bIsReportExternal As Boolean = False
    Private g_sSharedReportName As String = String.Empty
    Private g_sDocumentType As String = ""
    Private g_sExportPath As String = ""
    Private cClientType As String = "D"

#End Region

#Region "Property"
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
    Public Property IsReportExternal() As Boolean
        Get
            Return g_bIsReportExternal
        End Get
        Set(ByVal Value As Boolean)
            g_bIsReportExternal = Value
        End Set
    End Property
    Public Property DocumentType() As String
        Get
            Return g_sDocumentType
        End Get
        Set(ByVal Value As String)
            g_sDocumentType = Value
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
    Public Property ReportPath() As String
        Get
            Return cReportPath
        End Get
        Set(ByVal value As String)
            cReportPath = value
        End Set
    End Property
    Public Property ReportDataSet() As DataSet
        Get
            Return cReportDataSet
        End Get
        Set(ByVal value As DataSet)
            cReportDataSet = value
        End Set
    End Property
    Public Property Report() As Integer
        Get
            Return cReport
        End Get
        Set(ByVal Value As Integer)
            cReport = Value
        End Set
    End Property
    Public Property UserCode() As String
        Get
            Return cUserCode
        End Get
        Set(ByVal value As String)
            cUserCode = value
        End Set
    End Property
#End Region

#Region "Reports Printing"
    Friend Sub OpenRPLReport()
        Try
            Dim rpt As New ReportDocument
            Dim bLoadDatasource As Boolean = True

            If g_bIsReportExternal = False Then
                rpt = New rptRecpPaym
            Else
                rpt.Load(g_sSharedReportName)
            End If

            rpt.SetDataSource(cReportDataSet)
            rpt.DataDefinition.FormulaFields.Item("CompanyNameFormula").Text = "'" & cCompanyName & "'"
            rpt.DataDefinition.FormulaFields.Item("DOCTYPE").Text = "'" & g_sDocumentType & "'"

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
                RecpPaym_crViewer.ReportSource = rpt
                RecpPaym_crViewer.Show()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub RecpPaym_crViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecpPaym_crViewer.Load
        Try
            OpenRPLReport()

        Catch ex As Exception
            MsgBox("[Frm_Viewer].[frmViewer_Load]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region



End Class
