'' © Copyright © 2007-2019, Inecom Pte Ltd, All rights reserved.
'' =============================================================

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Public Class Hydac_FormViewer
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
    Friend WithEvents crViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.crViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'crViewer
        '
        Me.crViewer.ActiveViewIndex = -1
        Me.crViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crViewer.EnableDrillDown = False
        Me.crViewer.Location = New System.Drawing.Point(0, 0)
        Me.crViewer.Name = "crViewer"
        Me.crViewer.ReportSource = Nothing
        Me.crViewer.Size = New System.Drawing.Size(480, 437)
        Me.crViewer.TabIndex = 0
        '
        'frmViewer
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 437)
        Me.Controls.Add(Me.crViewer)
        Me.Name = "frmViewer"
        Me.Text = "Report Viewer"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CONST_ODBC_SERVER_NAME As String = "TEST_HANA_USER"

    Private cBankAccount As String = ""
    Private cBankDate As String = ""

    Private cCRMDateFr As String = ""
    Private cCRMDateTo As String = ""
    Private cServer As String = ""
    Private cDatabase As String = ""
    Private cAsAtDate As String = ""
    Private cUsername As String = ""
    Private cAgeBy As Integer
    Private cReportType As AgeingType
    Private cProjects As String = ""
    Private cReport As ReportCode
    Private cPeriod As Integer
    Private cServerName As String = ""
    Private cDBUser As String = ""
    Private cDBPassword As String = ""
    Private cHideLogo As Boolean
    Private cHideHeader As Boolean
    Private cIsBBF As Integer
    Private cIsSNP As Integer
    Private cIsGAT As Integer
    Private cIsHAS As Integer
    Private cIsHFN As Integer
    Private cReportName As ReportName
    Private cCompany As CompanyCode
    Private g_sSharedReportName As String = ""
    Private cIsLandscape As String = ""
    Private sARSOARunningDate As String = ""
    Private sAPSOARunningDate As String = ""
    Private sARAGERunningDate As String = ""
    Private sAPAGERunningDate As String = ""

    Private sBPCode As String = String.Empty
    Private sBPCodeFr As String = String.Empty
    Private sBPCodeTo As String = String.Empty
    Private sBPGrpFr As String = String.Empty
    Private sBPGrpTo As String = String.Empty
    Private sSlsFr As String = String.Empty
    Private sSlsTo As String = String.Empty
    Private sLocalCurr As String = String.Empty
    Private iSectionPageBreak As Integer = 0
    Private sAgingBy As String = String.Empty
    Private bIsExcel As Boolean = False

    Private bIsExport As Boolean = False
    Private sExportCardCode As String = String.Empty
    Private oExportType As CrystalDecisions.Shared.ExportFormatType = ExportFormatType.PortableDocFormat
    Private sExportPath As String = String.Empty
    Private sDatabaseServer As String = String.Empty
    Private sDatabaseName As String = String.Empty

    Private cDateFrom As String = ""
    Private cDateTo As String = ""
    Private cBPFrom As String = ""
    Private cBPTo As String = ""
    Private cReportPath As String = String.Empty
    Private cReportDataSet As DataSet
    Private cUserCode As String = ""
    Private cBusiness As String = ""

    Private cDocNum As String = ""
    Private cSeries As String = ""
    Private cDocEntry As String = ""
    Private cShowDetails As Boolean
    Private cShowTaxDate As String = ""
    Private g_bIsShared As Boolean
    Private g_sReportName As String = ""
    Private sExcelFilePath As String = String.Empty
    Private sBucketText As String() = New String(10) {}
    Private sBucketVal As Integer()

    Private iReportType As Integer
    Private dSmoothFactor As Decimal
    Private iNoOfWeek As Integer

    Private cFrDate As String = ""
    Private cToDate As String = ""
    Private cFrBP As String = ""
    Private cToBP As String = ""
    Private cFrProj As String = ""
    Private cToProj As String = ""
    Private cDataset As DataSet
    Private g_sExportPath As String = ""
    Private cClientType As String = ""


    Private CLOG_UserId As String = ""
    Private CLOG_CompanyName As String = ""
    Private CLOG_GenBy As String = ""
    Private CLOG_dtFrom As Date
    Private CLOG_dtTo As Date


#Region "Property Change Log"
    Public Property CL_CompanyNm() As String
        Get
            Return CLOG_CompanyName
        End Get
        Set(ByVal Value As String)
            CLOG_CompanyName = Value
        End Set
    End Property

    Public Property CL_UserID() As String
        Get
            Return CLOG_UserId
        End Get
        Set(ByVal Value As String)
            CLOG_UserId = Value
        End Set
    End Property

    Public Property CL_DateFrom() As Date
        Get
            Return CLOG_dtFrom
        End Get
        Set(ByVal Value As Date)
            CLOG_dtFrom = Value
        End Set
    End Property

    Public Property CL_DateTo() As Date
        Get
            Return CLOG_dtTo
        End Get
        Set(ByVal Value As Date)
            CLOG_dtTo = Value
        End Set
    End Property
    Public Property CL_GenBy() As String
        Get
            Return CLOG_GenBy
        End Get
        Set(ByVal Value As String)
            CLOG_GenBy = Value
        End Set
    End Property
#End Region

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
    Friend Property BankAccount() As String
        Get
            Return cBankAccount
        End Get
        Set(ByVal Value As String)
            cBankAccount = Value
        End Set
    End Property
    Friend Property BankDate() As String
        Get
            Return cBankDate
        End Get
        Set(ByVal Value As String)
            cBankDate = Value
        End Set
    End Property

    Friend Property CRMDateFr() As String
        Get
            Return cCRMDateFr
        End Get
        Set(ByVal Value As String)
            cCRMDateFr = Value
        End Set
    End Property
    Friend Property CRMDateTo() As String
        Get
            Return cCRMDateTo
        End Get
        Set(ByVal Value As String)
            cCRMDateTo = Value
        End Set
    End Property

    Friend Property ARSOARunningDate() As String
        Get
            Return sARSOARunningDate
        End Get
        Set(ByVal Value As String)
            sARSOARunningDate = Value
        End Set
    End Property
    Friend Property APSOARunningDate() As String
        Get
            Return sAPSOARunningDate
        End Get
        Set(ByVal Value As String)
            sAPSOARunningDate = Value
        End Set
    End Property
    Friend Property ARAGERunningDate() As String
        Get
            Return sARAGERunningDate
        End Get
        Set(ByVal Value As String)
            sARAGERunningDate = Value
        End Set
    End Property
    Friend Property APAGERunningDate() As String
        Get
            Return sAPAGERunningDate
        End Get
        Set(ByVal Value As String)
            sAPAGERunningDate = Value
        End Set
    End Property

    Friend Property Landscape() As String
        Get
            Return cIsLandscape
        End Get
        Set(ByVal Value As String)
            cIsLandscape = Value
        End Set
    End Property
    Friend Property Projects() As String
        Get
            Return cProjects
        End Get
        Set(ByVal Value As String)
            cProjects = Value
        End Set
    End Property
    Friend Property Server() As String
        Get
            Return cServerName
        End Get
        Set(ByVal Value As String)
            cServerName = Value
        End Set
    End Property
    Friend Property Database() As String
        Get
            Return cDatabase
        End Get
        Set(ByVal Value As String)
            cDatabase = Value
        End Set
    End Property
    Friend Property DBUsername() As String
        Get
            Return cDBUser
        End Get
        Set(ByVal Value As String)
            cDBUser = Value
        End Set
    End Property
    Friend Property DBPassword() As String
        Get
            Return cDBPassword
        End Get
        Set(ByVal Value As String)
            cDBPassword = Value
        End Set
    End Property
    Friend Property SmoothFactor() As Decimal
        Get
            Return dSmoothFactor
        End Get
        Set(ByVal Value As Decimal)
            dSmoothFactor = Value
        End Set
    End Property
    Friend Property NoOfWeek() As Integer
        Get
            Return iNoOfWeek
        End Get
        Set(ByVal Value As Integer)
            iNoOfWeek = Value
        End Set
    End Property

    Public Property Username() As String
        Get
            Return cUsername
        End Get
        Set(ByVal Value As String)
            cUsername = Value
        End Set
    End Property
    Public Property AsAtDate() As String
        Get
            Return cAsAtDate
        End Get
        Set(ByVal Value As String)
            cAsAtDate = Value
        End Set
    End Property
    Public Property AgeBy() As Integer
        Get
            Return cAgeBy
        End Get
        Set(ByVal Value As Integer)
            cAgeBy = Value
        End Set
    End Property
    Public Property ReportType() As Integer
        Get
            Return cReportType
        End Get
        Set(ByVal Value As Integer)
            cReportType = Value
        End Set
    End Property

    Public Property Dataset() As DataSet
        Get
            Return cDataset
        End Get
        Set(ByVal Value As DataSet)
            cDataset = Value
        End Set
    End Property
    Public Property FromDate() As String
        Get
            Return cFrDate
        End Get
        Set(ByVal Value As String)
            cFrDate = Value
        End Set
    End Property
    Public Property ToDate() As String
        Get
            Return cToDate
        End Get
        Set(ByVal Value As String)
            cToDate = Value
        End Set
    End Property
    Public Property FromBP() As String
        Get
            Return cFrBP
        End Get
        Set(ByVal Value As String)
            cFrBP = Value
        End Set
    End Property
    Public Property ToBP() As String
        Get
            Return cToBP
        End Get
        Set(ByVal Value As String)
            cToBP = Value
        End Set
    End Property
    Public Property FromProj() As String
        Get
            Return cFrProj
        End Get
        Set(ByVal Value As String)
            cFrProj = Value
        End Set
    End Property
    Public Property ToProj() As String
        Get
            Return cToProj
        End Get
        Set(ByVal Value As String)
            cToProj = Value
        End Set
    End Property
    Public Property LineBusiness() As String
        Get
            Return cBusiness
        End Get
        Set(ByVal Value As String)
            cBusiness = Value
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
    Public Property ReportName() As ReportName
        Get
            Return cReportName
        End Get
        Set(ByVal Value As ReportName)
            cReportName = Value
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
    Public Property Period() As Integer
        Get
            Return cPeriod
        End Get
        Set(ByVal Value As Integer)
            cPeriod = Value
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
    Public Property HideLogo() As Boolean
        Get
            Return cHideLogo
        End Get
        Set(ByVal Value As Boolean)
            cHideLogo = Value
        End Set
    End Property
    Public Property HideHeader() As Boolean
        Get
            Return cHideHeader
        End Get
        Set(ByVal Value As Boolean)
            cHideHeader = Value
        End Set
    End Property
    Public Property IsBBF() As Integer
        Get
            Return cIsBBF
        End Get
        Set(ByVal Value As Integer)
            cIsBBF = Value
        End Set
    End Property
    Public Property IsSNP() As Integer
        Get
            Return cIsSNP
        End Get
        Set(ByVal Value As Integer)
            cIsSNP = Value
        End Set
    End Property
    Public Property IsGAT() As Integer
        Get
            Return cIsGAT
        End Get
        Set(ByVal Value As Integer)
            cIsGAT = Value
        End Set
    End Property
    Public Property IsHAS() As Integer
        Get
            Return cIsHAS
        End Get
        Set(ByVal Value As Integer)
            cIsHAS = Value
        End Set
    End Property
    Public Property IsHFN() As Integer
        Get
            Return cIsHFN
        End Get
        Set(ByVal Value As Integer)
            cIsHFN = Value
        End Set
    End Property
    Public Property CompanySOA() As CompanyCode
        Get
            Return cCompany
        End Get
        Set(ByVal Value As CompanyCode)
            cCompany = Value
        End Set
    End Property
    Public Property BPCode() As String
        Get
            Return sBPCode
        End Get
        Set(ByVal Value As String)
            sBPCode = Value
        End Set
    End Property
    Public Property BPCodeFr() As String
        Get
            Return sBPCodeFr
        End Get
        Set(ByVal Value As String)
            sBPCodeFr = Value
        End Set
    End Property
    Public Property BPCodeTo() As String
        Get
            Return sBPCodeTo
        End Get
        Set(ByVal Value As String)
            sBPCodeTo = Value
        End Set
    End Property
    Public Property BPGroupFr() As String
        Get
            Return sBPGrpFr
        End Get
        Set(ByVal Value As String)
            sBPGrpFr = Value
        End Set
    End Property
    Public Property BPGroupTo() As String
        Get
            Return sBPGrpTo
        End Get
        Set(ByVal Value As String)
            sBPGrpTo = Value
        End Set
    End Property
    Public Property SalesEmployeeFr() As String
        Get
            Return sSlsFr
        End Get
        Set(ByVal Value As String)
            sSlsFr = Value
        End Set
    End Property
    Public Property SalesEmployeeTo() As String
        Get
            Return sSlsTo
        End Get
        Set(ByVal Value As String)
            sSlsTo = Value
        End Set
    End Property
    Public Property LocalCurrency() As String
        Get
            Return sLocalCurr
        End Get
        Set(ByVal Value As String)
            sLocalCurr = Value
        End Set
    End Property
    Public Property SectionPageBreak() As Integer
        Get
            Return iSectionPageBreak
        End Get
        Set(ByVal Value As Integer)
            iSectionPageBreak = Value
        End Set
    End Property
    Public Property AgingBy() As String
        Get
            Return sAgingBy
        End Get
        Set(ByVal Value As String)
            sAgingBy = Value
        End Set
    End Property

    Public Property IsExport() As Boolean
        Get
            Return bIsExport
        End Get
        Set(ByVal value As Boolean)
            bIsExport = value
        End Set
    End Property
    Public Property ExportCustomerCode() As String
        Get
            Return sExportCardCode
        End Get
        Set(ByVal value As String)
            sExportCardCode = value
        End Set
    End Property
    Public Property CrystalReportExportType() As CrystalDecisions.Shared.ExportFormatType
        Get
            Return oExportType
        End Get
        Set(ByVal value As CrystalDecisions.Shared.ExportFormatType)
            oExportType = value
        End Set
    End Property
    Public Property CrystalReportExportPath() As String
        Get
            Return sExportPath
        End Get
        Set(ByVal value As String)
            sExportPath = value
        End Set
    End Property
    Public Property DatabaseServer() As String
        Get
            Return sDatabaseServer

        End Get
        Set(ByVal value As String)
            sDatabaseServer = value
        End Set
    End Property
    Public Property DatabaseName() As String
        Get
            Return sDatabaseName
        End Get
        Set(ByVal value As String)
            sDatabaseName = value
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
    Public Property UserCode() As String
        Get
            Return cUserCode
        End Get
        Set(ByVal value As String)
            cUserCode = value
        End Set
    End Property
    Public Property ParamDateFrom() As String
        Get
            Return cDateFrom
        End Get
        Set(ByVal value As String)
            cDateFrom = value
        End Set
    End Property
    Public Property ParamDateTo() As String
        Get
            Return cDateTo
        End Get
        Set(ByVal value As String)
            cDateTo = value
        End Set
    End Property
    Public Property ParamBPFrom() As String
        Get
            Return cBPFrom
        End Get
        Set(ByVal value As String)
            cBPFrom = value
        End Set
    End Property
    Public Property ParamBPTo() As String
        Get
            Return cBPTo
        End Get
        Set(ByVal value As String)
            cBPTo = value
        End Set
    End Property

    Public Property DocNum() As String
        Get
            Return cDocNum
        End Get
        Set(ByVal Value As String)
            cDocNum = Value
        End Set
    End Property
    Public Property Series() As String
        Get
            Return cSeries
        End Get
        Set(ByVal Value As String)
            cSeries = Value
        End Set
    End Property
    Public Property DocEntry() As String
        Get
            Return cDocEntry
        End Get
        Set(ByVal Value As String)
            cDocEntry = Value
        End Set
    End Property
    Public Property ReportNamePV() As String
        Get
            Return g_sReportName
        End Get
        Set(ByVal Value As String)
            g_sReportName = Value
        End Set
    End Property
    Public Property ShowDetails() As Boolean
        Get
            Return cShowDetails
        End Get
        Set(ByVal Value As Boolean)
            cShowDetails = Value
        End Set
    End Property
    Public Property ShowTaxDate() As String
        Get
            Return cShowTaxDate
        End Get
        Set(ByVal Value As String)
            cShowTaxDate = Value
        End Set
    End Property
    Public Property IsShared() As Boolean
        Get
            Return g_bIsShared
        End Get
        Set(ByVal Value As Boolean)
            g_bIsShared = Value
        End Set
    End Property
    Public Property IsExcel() As Boolean
        Get
            Return bIsExcel
        End Get
        Set(ByVal Value As Boolean)
            bIsExcel = Value
        End Set
    End Property
    Friend Property ExcelFilePath() As String
        Get
            Return sExcelFilePath
        End Get
        Set(ByVal Value As String)
            sExcelFilePath = Value
        End Set
    End Property
    Friend Property BucketText() As String()
        Get
            Return sBucketText
        End Get
        Set(ByVal Value As String())
            sBucketText = Value
        End Set
    End Property
    Friend Property BucketValue() As Integer()
        Get
            Return sBucketVal
        End Get
        Set(ByVal Value As Integer())
            sBucketVal = Value
        End Set
    End Property

#Region "PVRange Area"
    Private sDocNumS As String = String.Empty
    Private sDocNumE As String = String.Empty
    Private sBPCodeS As String = String.Empty
    Private sBPCodeE As String = String.Empty
    Private sDocDateS As String = String.Empty
    Private sDocDateE As String = String.Empty
    Private dtDocDateS As DateTime
    Private dtDocDateE As DateTime
    Private iIsIncludeCancel As Integer = 0
    Private myPaymentVoucherDocType As PaymentVoucherRangeDocTypes = PaymentVoucherRangeDocTypes.Supplier

    Public Property DocNumStart() As String
        Get
            Return sDocNumS
        End Get
        Set(ByVal Value As String)
            sDocNumS = Value
        End Set
    End Property
    Public Property DocNumEnd() As String
        Get
            Return sDocNumE
        End Get
        Set(ByVal Value As String)
            sDocNumE = Value
        End Set
    End Property
    Public Property BPCodeStart() As String
        Get
            Return sBPCodeS
        End Get
        Set(ByVal Value As String)
            sBPCodeS = Value
        End Set
    End Property
    Public Property BPCodeEnd() As String
        Get
            Return sBPCodeE
        End Get
        Set(ByVal Value As String)
            sBPCodeE = Value
        End Set
    End Property

    Public Property DocDateSStart() As String
        Get
            Return sDocDateS
        End Get
        Set(ByVal Value As String)
            sDocDateS = Value
        End Set
    End Property
    Public Property DocDateSEnd() As String
        Get
            Return sDocDateE
        End Get
        Set(ByVal Value As String)
            sDocDateE = Value
        End Set
    End Property
    Public Property DocDateStart() As DateTime
        Get
            Return dtDocDateS
        End Get
        Set(ByVal Value As DateTime)
            dtDocDateS = Value
        End Set

    End Property
    Public Property DocDateEnd() As DateTime
        Get
            Return dtDocDateE
        End Get
        Set(ByVal Value As DateTime)
            dtDocDateE = Value
        End Set
    End Property

    Public Property IsIncludeCancel() As Integer
        Get
            Return iIsIncludeCancel
        End Get
        Set(ByVal Value As Integer)
            iIsIncludeCancel = Value
        End Set
    End Property
    Public Property PVRangeDocType() As PaymentVoucherRangeDocTypes
        Get
            Return myPaymentVoucherDocType
        End Get
        Set(ByVal Value As PaymentVoucherRangeDocTypes)
            myPaymentVoucherDocType = Value
        End Set
    End Property
    Private Function GetPVRangeSelectionFormula() As String
        Dim sDateFormat As String = " AND  {0} Date({1},{2},{3}) "
        Dim sDocNumFormat As String = " AND  {0} {1} "
        Dim sBPFormat As String = " AND  {0} '{1}' "

        Dim sOutput As String = String.Empty
        Select Case myPaymentVoucherDocType
            Case PaymentVoucherRangeDocTypes.All
                sOutput = "1 = 1 "
            Case PaymentVoucherRangeDocTypes.Account
                sOutput = " {OVPM.DocType} = 'A' "
            Case PaymentVoucherRangeDocTypes.Customer
                sOutput = " {OVPM.DocType} = 'C' "
            Case PaymentVoucherRangeDocTypes.Supplier
                sOutput = " {OVPM.DocType} = 'S' "
        End Select

        If sDocDateS.Length > 0 Then
            sOutput = sOutput & String.Format(sDateFormat, "{OVPM.DocDate} >=", dtDocDateS.Year.ToString("0000"), dtDocDateS.Month.ToString("00"), dtDocDateS.Day.ToString("00"))
        End If
        If sDocDateE.Length > 0 Then
            sOutput = sOutput & String.Format(sDateFormat, "{OVPM.DocDate} <=", dtDocDateE.Year.ToString("0000"), dtDocDateE.Month.ToString("00"), dtDocDateE.Day.ToString("00"))
        End If

        If (sDocNumS.Length > 0) Then
            sOutput = sOutput & String.Format(sDocNumFormat, "{OVPM.DocNum} >=", sDocNumS)
        End If

        If (sDocNumE.Length > 0) Then
            sOutput = sOutput & String.Format(sDocNumFormat, "{OVPM.DocNum} <=", sDocNumE)
        End If

        If (sBPCodeS.Length > 0) Then
            sOutput = sOutput & String.Format(sBPFormat, "{OVPM.CardCode} >=", sBPCodeS)
        End If

        If (sBPCodeE.Length > 0) Then
            sOutput = sOutput & String.Format(sBPFormat, "{OVPM.CardCode} <=", sBPCodeE)
        End If

        If (iIsIncludeCancel = 0) Then
            sOutput = sOutput & " AND {OVPM.Canceled} = 'N' "
        End If
        Return sOutput
    End Function
#End Region

    Friend Sub OpenAgingReport_HANA_CRM()
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        crConnectionInfo.ServerName = CONST_ODBC_SERVER_NAME
        crConnectionInfo.DatabaseName = oCompany.CompanyDB
        crConnectionInfo.UserID = global_DBUsername
        crConnectionInfo.Password = global_DBPassword

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New rptARAgeing_HANA_CRM
        End If

        With rpt
            ' Main report --------------------------------------------------------------------------------
            CrTables = .Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next
            ' Sub Report for Company Details -------------------------------------------------------------
            subreport = .OpenSubreport("CH.rpt")
            CrTables = subreport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            Dim bIsUserDefinedRange As Boolean = False
            Dim iCount As Integer = 1
            Dim sFormat As String = "Bucket{0}Text"
            Dim sTemp As String = String.Format(sFormat, 1)

            Try
                .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                bIsUserDefinedRange = True
            Catch ex As Exception
                bIsUserDefinedRange = False
            End Try

            If bIsUserDefinedRange Then
                sFormat = "Bucket{0}Text"
                For iCount = 1 To 5 Step 1
                    sTemp = String.Format(sFormat, iCount)
                    .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                Next

                sFormat = "Bucket{0}Val"
                For iCount = 1 To 5 Step 1
                    sTemp = String.Format(sFormat, iCount)
                    .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                Next
            End If

            ' Sub Report (SUBCRM) --------------------------------------------
            subreport = .OpenSubreport("SUBCRM")
            CrTables = subreport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next
            subreport.RecordSelectionFormula &= " AND {OCLG.CntctDate} >= DateValue(" & cCRMDateFr.Substring(0, 4) & "," & cCRMDateFr.Substring(4, 2) & "," & cCRMDateFr.Substring(6, 2) & ") "
            subreport.RecordSelectionFormula &= " AND {OCLG.CntctDate} <= DateValue(" & cCRMDateTo.Substring(0, 4) & "," & cCRMDateTo.Substring(4, 2) & "," & cCRMDateTo.Substring(6, 2) & ") "
            subreport.RecordSelectionFormula &= " AND {OCLT.Name} = 'Collection' "

            ' Sub Report Company Details -----------------------------------------------------------------
            Dim StatementDate As String = "DateValue(" & cAsAtDate.Substring(0, 4) & "," & cAsAtDate.Substring(4, 2) & "," & cAsAtDate.Substring(6, 2) & ")"

            subreport = .OpenSubreport("PageHeader.rpt")
            subreport.DataDefinition.FormulaFields.Item("CompanyId").Text = SetStringIntoCrystalFormula(oCompany.CompanyName)
            subreport.DataDefinition.FormulaFields.Item("BPCode").Text = SetStringIntoCrystalFormula(sBPCode)
            subreport.DataDefinition.FormulaFields.Item("BPCodeFr").Text = SetStringIntoCrystalFormula(sBPCodeFr)
            subreport.DataDefinition.FormulaFields.Item("BPCodeTo").Text = SetStringIntoCrystalFormula(sBPCodeTo)
            subreport.DataDefinition.FormulaFields.Item("BPGrpFr").Text = SetStringIntoCrystalFormula(sBPGrpFr)
            subreport.DataDefinition.FormulaFields.Item("BPGrpTo").Text = SetStringIntoCrystalFormula(sBPGrpTo)
            subreport.DataDefinition.FormulaFields.Item("SlsFr").Text = SetStringIntoCrystalFormula(sSlsFr)
            subreport.DataDefinition.FormulaFields.Item("SlsTo").Text = SetStringIntoCrystalFormula(sSlsTo)
            subreport.DataDefinition.FormulaFields.Item("AsAtDate").Text = StatementDate
            subreport.DataDefinition.FormulaFields.Item("AgingBy").Text = SetStringIntoCrystalFormula(sAgingBy)
            subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Details Report with CRM Notes")
            If (iSectionPageBreak = 0) Then
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("No")
            Else
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("Yes")
            End If


            ' Main Report Parameter ----------------------------------------------------------------------
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
            .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
            .DataDefinition.FormulaFields.Item("PageBreak").Text = iSectionPageBreak
            .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
            .Refresh()
        End With

        If (bIsExcel) Then
            Dim sPath As String = String.Empty
            sPath = sExcelFilePath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
            Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rpt.ExportOptions.FormatOptions = objExcelOptions
            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
            System.Diagnostics.Process.Start(sPath)
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub

    Friend Sub OpenAgingReport()
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        crConnectionInfo.ServerName = CONST_ODBC_SERVER_NAME
        crConnectionInfo.DatabaseName = oCompany.CompanyDB
        crConnectionInfo.UserID = global_DBUsername
        crConnectionInfo.Password = global_DBPassword

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            Select Case ReportName
                Case ReportName.ARAging7B_Details
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing7B_HANA
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary7B_HANA
                    End Select

                Case ReportName.ARAging6B_Details
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing6B_HANA
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary6B_HANA
                    End Select
                Case Else
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing_HANA
                        Case AgeingType.APAgeing
                            rpt = New rptAPAgeing_HANA
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary_HANA
                        Case AgeingType.APAgeingSummary
                            rpt = New rptAPAgeingSummary_HANA
                    End Select
            End Select
        End If

        With rpt
            ' Main report --------------------------------------------------------------------------------
            CrTables = .Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next
            ' Sub Report for Company Details -------------------------------------------------------------
            subreport = .OpenSubreport("CH.rpt")
            CrTables = subreport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            Dim bIsUserDefinedRange As Boolean = False
            Dim iCount As Integer = 1
            Dim sFormat As String = "Bucket{0}Text"
            Dim sTemp As String = String.Format(sFormat, 1)

            Try
                .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                bIsUserDefinedRange = True
            Catch ex As Exception
                bIsUserDefinedRange = False
            End Try

            If bIsUserDefinedRange Then
                Select Case ReportName
                    Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case Else
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next
                End Select
            End If

            ' Sub Report (Only for AP Summary and AR Summary) --------------------------------------------
            If cReportType = AgeingType.APAgeingSummary Or cReportType = AgeingType.ARAgeingSummary Then
                subreport = .OpenSubreport("Summary")
                CrTables = subreport.Database.Tables
                For Each CrTable In CrTables
                    crtableLogoninfo = CrTable.LogOnInfo
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
                    CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                Next

                subreport.DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
                subreport.DataDefinition.FormulaFields.Item("LocalCurr").Text = SetStringIntoCrystalFormula(sLocalCurr)

                If bIsUserDefinedRange Then
                    sFormat = "Bucket{0}Val"
                    Select Case ReportName
                        Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                            For iCount = 1 To 6 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                            For iCount = 1 To 7 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case Else
                            For iCount = 1 To 5 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                    End Select
                End If
            End If

            ' Sub Report Company Details -----------------------------------------------------------------
            Dim StatementDate As String = "DateValue(" & cAsAtDate.Substring(0, 4) & "," & cAsAtDate.Substring(4, 2) & "," & cAsAtDate.Substring(6, 2) & ")"
            subreport = .OpenSubreport("PageHeader.rpt")
            subreport.DataDefinition.FormulaFields.Item("CompanyId").Text = SetStringIntoCrystalFormula(oCompany.CompanyName)
            subreport.DataDefinition.FormulaFields.Item("BPCode").Text = SetStringIntoCrystalFormula(sBPCode)
            subreport.DataDefinition.FormulaFields.Item("BPCodeFr").Text = SetStringIntoCrystalFormula(sBPCodeFr)
            subreport.DataDefinition.FormulaFields.Item("BPCodeTo").Text = SetStringIntoCrystalFormula(sBPCodeTo)
            subreport.DataDefinition.FormulaFields.Item("BPGrpFr").Text = SetStringIntoCrystalFormula(sBPGrpFr)
            subreport.DataDefinition.FormulaFields.Item("BPGrpTo").Text = SetStringIntoCrystalFormula(sBPGrpTo)
            subreport.DataDefinition.FormulaFields.Item("SlsFr").Text = SetStringIntoCrystalFormula(sSlsFr)
            subreport.DataDefinition.FormulaFields.Item("SlsTo").Text = SetStringIntoCrystalFormula(sSlsTo)
            subreport.DataDefinition.FormulaFields.Item("AsAtDate").Text = StatementDate
            subreport.DataDefinition.FormulaFields.Item("AgingBy").Text = SetStringIntoCrystalFormula(sAgingBy)
            If (iSectionPageBreak = 0) Then
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("No")
            Else
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("Yes")
            End If

            Select Case cReportType
                Case AgeingType.APAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AP Ageing Details")
                Case AgeingType.APAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AP Ageing Summary")
                Case AgeingType.ARAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Details")
                Case AgeingType.ARAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Summary")
            End Select

            ' Main Report Parameter ----------------------------------------------------------------------
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
            .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
            .DataDefinition.FormulaFields.Item("PageBreak").Text = iSectionPageBreak

            If ReportName = ReportName.ARAging6B_Details Then
                .DataDefinition.FormulaFields.Item("IsLocalCurr").Text = 1
            End If

            Select Case cReportType
                Case AgeingType.ARAgeing
                    .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
                Case AgeingType.ARAgeingSummary
                    .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
                Case AgeingType.APAgeing
                    .RecordSelectionFormula = "{_NCM_AP_AGEING.USERNAME}='" & sAPAGERunningDate & "'"
                Case AgeingType.APAgeingSummary
                    .RecordSelectionFormula = "{_NCM_AP_AGEING.USERNAME}='" & sAPAGERunningDate & "'"
            End Select
            .Refresh()
        End With

        If (bIsExcel) Then
            Dim sPath As String = String.Empty
            sPath = sExcelFilePath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
            Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rpt.ExportOptions.FormatOptions = objExcelOptions
            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
            System.Diagnostics.Process.Start(sPath)
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub
    Friend Sub OpenAgingReportProject_HANA()
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        crConnectionInfo.ServerName = CONST_ODBC_SERVER_NAME
        crConnectionInfo.DatabaseName = oCompany.CompanyDB
        crConnectionInfo.UserID = global_DBUsername
        crConnectionInfo.Password = global_DBPassword

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            Select Case ReportName
                Case ReportName.ARAging_Details_Proj
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing_PROJ_HANA
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary_PROJ_HANA
                    End Select
                Case ReportName.APAging_Details_Proj
                    Select Case cReportType
                        Case AgeingType.APAgeing
                            rpt = New rptAPAgeing_PROJ_HANA
                        Case AgeingType.APAgeingSummary
                            rpt = New rptAPAgeingSummary_PROJ_HANA
                    End Select
            End Select
        End If

        With rpt
            ' Main report --------------------------------------------------------------------------------
            CrTables = .Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next
            ' Sub Report for Company Details -------------------------------------------------------------
            subreport = .OpenSubreport("CH.rpt")
            CrTables = subreport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            Dim bIsUserDefinedRange As Boolean = False
            Dim iCount As Integer = 1
            Dim sFormat As String = "Bucket{0}Text"
            Dim sTemp As String = String.Format(sFormat, 1)

            Try
                .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                bIsUserDefinedRange = True
            Catch ex As Exception
                bIsUserDefinedRange = False
            End Try

            If bIsUserDefinedRange Then
                Select Case ReportName
                    Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case Else
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next
                End Select
            End If

            ' Sub Report (Only for AP Summary and AR Summary) --------------------------------------------
            If cReportType = AgeingType.APAgeingSummary Or cReportType = AgeingType.ARAgeingSummary Then
                subreport = .OpenSubreport("Summary")
                CrTables = subreport.Database.Tables
                For Each CrTable In CrTables
                    crtableLogoninfo = CrTable.LogOnInfo
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
                    CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                Next

                subreport.DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
                subreport.DataDefinition.FormulaFields.Item("LocalCurr").Text = SetStringIntoCrystalFormula(sLocalCurr)

                If bIsUserDefinedRange Then
                    sFormat = "Bucket{0}Val"
                    Select Case ReportName
                        Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                            For iCount = 1 To 6 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                            For iCount = 1 To 7 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case Else
                            For iCount = 1 To 5 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                    End Select
                End If
            End If

            ' Sub Report Company Details -----------------------------------------------------------------
            Dim StatementDate As String = "DateValue(" & cAsAtDate.Substring(0, 4) & "," & cAsAtDate.Substring(4, 2) & "," & cAsAtDate.Substring(6, 2) & ")"
            subreport = .OpenSubreport("PageHeader.rpt")
            subreport.DataDefinition.FormulaFields.Item("CompanyId").Text = SetStringIntoCrystalFormula(oCompany.CompanyName)
            subreport.DataDefinition.FormulaFields.Item("BPCode").Text = SetStringIntoCrystalFormula(sBPCode)
            subreport.DataDefinition.FormulaFields.Item("BPCodeFr").Text = SetStringIntoCrystalFormula(sBPCodeFr)
            subreport.DataDefinition.FormulaFields.Item("BPCodeTo").Text = SetStringIntoCrystalFormula(sBPCodeTo)
            subreport.DataDefinition.FormulaFields.Item("BPGrpFr").Text = SetStringIntoCrystalFormula(sBPGrpFr)
            subreport.DataDefinition.FormulaFields.Item("BPGrpTo").Text = SetStringIntoCrystalFormula(sBPGrpTo)
            subreport.DataDefinition.FormulaFields.Item("SlsFr").Text = SetStringIntoCrystalFormula(sSlsFr)
            subreport.DataDefinition.FormulaFields.Item("SlsTo").Text = SetStringIntoCrystalFormula(sSlsTo)
            subreport.DataDefinition.FormulaFields.Item("ProjFr").Text = SetStringIntoCrystalFormula(cFrProj)
            subreport.DataDefinition.FormulaFields.Item("ProjTo").Text = SetStringIntoCrystalFormula(cToProj)
            subreport.DataDefinition.FormulaFields.Item("AsAtDate").Text = StatementDate
            subreport.DataDefinition.FormulaFields.Item("AgingBy").Text = SetStringIntoCrystalFormula(sAgingBy)
            If (iSectionPageBreak = 0) Then
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("No")
            Else
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("Yes")
            End If

            Select Case cReportType
                Case AgeingType.APAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = "'AP Ageing Details With Project'"
                Case AgeingType.APAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = "'AP Ageing Summary With Project'"
                Case AgeingType.ARAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = "'AR Ageing Details With Project'"
                Case AgeingType.ARAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = "'AR Ageing Summary With Project'"
            End Select

            ' Main Report Parameter ----------------------------------------------------------------------
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
            .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
            .DataDefinition.FormulaFields.Item("PageBreak").Text = iSectionPageBreak

            If ReportName = ReportName.ARAging6B_Details Then
                .DataDefinition.FormulaFields.Item("IsLocalCurr").Text = 1
            End If

            Select Case cReportType
                Case AgeingType.ARAgeing, AgeingType.ARAgeingSummary
                    Select Case cProjects.Trim.Length
                        Case Is <= 0
                            .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME} = '" & sARAGERunningDate & "'"
                        Case Is > 0
                            .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME} = '" & sARAGERunningDate & "' AND {_NCM_AR_AGEING.PROJECT} IN [" & cProjects & "] "
                    End Select
                Case AgeingType.APAgeing, AgeingType.APAgeingSummary
                    Select Case cProjects.Trim.Length
                        Case Is <= 0
                            .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME} = '" & sAPAGERunningDate & "'"
                        Case Is > 0
                            .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME} = '" & sAPAGERunningDate & "' AND {_NCM_AR_AGEING.PROJECT} IN [" & cProjects & "] "
                    End Select
            End Select
            .Refresh()
        End With

        If (bIsExcel) Then
            Dim sPath As String = String.Empty
            sPath = sExcelFilePath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
            Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rpt.ExportOptions.FormatOptions = objExcelOptions
            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
            System.Diagnostics.Process.Start(sPath)
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub

    Friend Sub OPEN_HANADS_AGEING_CRM()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New rptARAgeing_CRM_HANADS
        End If

        With rpt
            .SetDataSource(cDataset)

            Dim bIsUserDefinedRange As Boolean = False
            Dim iCount As Integer = 1
            Dim sFormat As String = "Bucket{0}Text"
            Dim sTemp As String = String.Format(sFormat, 1)

            Try
                .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                bIsUserDefinedRange = True
            Catch ex As Exception
                bIsUserDefinedRange = False
            End Try

            If bIsUserDefinedRange Then
                sFormat = "Bucket{0}Text"
                For iCount = 1 To 5 Step 1
                    sTemp = String.Format(sFormat, iCount)
                    .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                Next

                sFormat = "Bucket{0}Val"
                For iCount = 1 To 5 Step 1
                    sTemp = String.Format(sFormat, iCount)
                    .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                Next
            End If

            ' Sub Report (SUBCRM) --------------------------------------------
            subreport = .OpenSubreport("SUBCRM")
            subreport.RecordSelectionFormula &= " AND {OCLG.CntctDate} >= DateValue(" & cCRMDateFr.Substring(0, 4) & "," & cCRMDateFr.Substring(4, 2) & "," & cCRMDateFr.Substring(6, 2) & ") "
            subreport.RecordSelectionFormula &= " AND {OCLG.CntctDate} <= DateValue(" & cCRMDateTo.Substring(0, 4) & "," & cCRMDateTo.Substring(4, 2) & "," & cCRMDateTo.Substring(6, 2) & ") "
            subreport.RecordSelectionFormula &= " AND {OCLT.Name} = 'Collection' "

            ' Sub Report Company Details -----------------------------------------------------------------
            Dim StatementDate As String = "DateValue(" & cAsAtDate.Substring(0, 4) & "," & cAsAtDate.Substring(4, 2) & "," & cAsAtDate.Substring(6, 2) & ")"

            subreport = .OpenSubreport("PageHeader.rpt")
            subreport.DataDefinition.FormulaFields.Item("CompanyId").Text = SetStringIntoCrystalFormula(oCompany.CompanyName)
            subreport.DataDefinition.FormulaFields.Item("BPCode").Text = SetStringIntoCrystalFormula(sBPCode)
            subreport.DataDefinition.FormulaFields.Item("BPCodeFr").Text = SetStringIntoCrystalFormula(sBPCodeFr)
            subreport.DataDefinition.FormulaFields.Item("BPCodeTo").Text = SetStringIntoCrystalFormula(sBPCodeTo)
            subreport.DataDefinition.FormulaFields.Item("BPGrpFr").Text = SetStringIntoCrystalFormula(sBPGrpFr)
            subreport.DataDefinition.FormulaFields.Item("BPGrpTo").Text = SetStringIntoCrystalFormula(sBPGrpTo)
            subreport.DataDefinition.FormulaFields.Item("SlsFr").Text = SetStringIntoCrystalFormula(sSlsFr)
            subreport.DataDefinition.FormulaFields.Item("SlsTo").Text = SetStringIntoCrystalFormula(sSlsTo)
            subreport.DataDefinition.FormulaFields.Item("AsAtDate").Text = StatementDate
            subreport.DataDefinition.FormulaFields.Item("AgingBy").Text = SetStringIntoCrystalFormula(sAgingBy)
            subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Details Report with CRM Notes")
            If (iSectionPageBreak = 0) Then
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("No")
            Else
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("Yes")
            End If

            ' =====================
            ' Main Report Parameter ----------------------------------------------------------------------
            ' =====================

            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
            .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
            .DataDefinition.FormulaFields.Item("PageBreak").Text = iSectionPageBreak
            .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
            .Refresh()
        End With

        If (bIsExcel) Then
            Dim sPath As String = String.Empty
            sPath = sExcelFilePath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
            Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
            objExcelOptions.ExcelUseConstantColumnWidth = False
            rpt.ExportOptions.FormatOptions = objExcelOptions
            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
            System.Diagnostics.Process.Start(sPath)
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub
    Friend Sub OPEN_AGEING_REPORT_6_AND_7_BUCKETS()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            Select Case ReportName
                Case ReportName.ARAging7B_Details
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing7B_HANADS
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary7B_HANADS
                    End Select

                Case ReportName.ARAging6B_Details
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing6B_HANADS
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary6B_HANADS
                    End Select
                Case Else
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing_HANADS
                        Case AgeingType.APAgeing
                            rpt = New rptAPAgeing_HANADS
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary_HANADS
                        Case AgeingType.APAgeingSummary
                            rpt = New rptAPAgeingSummary_HANADS
                    End Select
            End Select
        End If

        With rpt
            Dim bIsUserDefinedRange As Boolean = False
            Dim iCount As Integer = 1
            Dim sFormat As String = "Bucket{0}Text"
            Dim sTemp As String = String.Format(sFormat, 1)

            Try
                .SetDataSource(cDataset)
                .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                bIsUserDefinedRange = True
            Catch ex As Exception
                bIsUserDefinedRange = False
            End Try

            ' ===================================================================================

            If bIsUserDefinedRange Then
                Select Case ReportName
                    Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case Else
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next
                End Select
            End If

            ' Sub Report (Only for AP Summary and AR Summary) --------------------------------------------
            If cReportType = AgeingType.APAgeingSummary Or cReportType = AgeingType.ARAgeingSummary Then
                subreport = .OpenSubreport("Summary")
                subreport.DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
                subreport.DataDefinition.FormulaFields.Item("LocalCurr").Text = SetStringIntoCrystalFormula(sLocalCurr)

                If bIsUserDefinedRange Then
                    sFormat = "Bucket{0}Val"
                    Select Case ReportName
                        Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                            For iCount = 1 To 6 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                            For iCount = 1 To 7 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case Else
                            For iCount = 1 To 5 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                    End Select
                End If
            End If

            ' Sub Report Company Details -----------------------------------------------------------------
            Dim StatementDate As String = "DateValue(" & cAsAtDate.Substring(0, 4) & "," & cAsAtDate.Substring(4, 2) & "," & cAsAtDate.Substring(6, 2) & ")"

            subreport = .OpenSubreport("PageHeader.rpt")
            subreport.DataDefinition.FormulaFields.Item("CompanyId").Text = SetStringIntoCrystalFormula(oCompany.CompanyName)
            subreport.DataDefinition.FormulaFields.Item("BPCode").Text = SetStringIntoCrystalFormula(sBPCode)
            subreport.DataDefinition.FormulaFields.Item("BPCodeFr").Text = SetStringIntoCrystalFormula(sBPCodeFr)
            subreport.DataDefinition.FormulaFields.Item("BPCodeTo").Text = SetStringIntoCrystalFormula(sBPCodeTo)
            subreport.DataDefinition.FormulaFields.Item("BPGrpFr").Text = SetStringIntoCrystalFormula(sBPGrpFr)
            subreport.DataDefinition.FormulaFields.Item("BPGrpTo").Text = SetStringIntoCrystalFormula(sBPGrpTo)
            subreport.DataDefinition.FormulaFields.Item("SlsFr").Text = SetStringIntoCrystalFormula(sSlsFr)
            subreport.DataDefinition.FormulaFields.Item("SlsTo").Text = SetStringIntoCrystalFormula(sSlsTo)
            subreport.DataDefinition.FormulaFields.Item("AsAtDate").Text = StatementDate
            subreport.DataDefinition.FormulaFields.Item("AgingBy").Text = SetStringIntoCrystalFormula(sAgingBy)
            If (iSectionPageBreak = 0) Then
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("No")
            Else
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("Yes")
            End If

            Select Case cReportType
                Case AgeingType.APAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AP Ageing Details")
                Case AgeingType.APAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AP Ageing Summary")
                Case AgeingType.ARAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Details")
                Case AgeingType.ARAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Summary")
            End Select

            ' Main Report Parameter ----------------------------------------------------------------------
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
            .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
            .DataDefinition.FormulaFields.Item("PageBreak").Text = iSectionPageBreak

            If ReportName = ReportName.ARAging6B_Details Then
                Try
                    .DataDefinition.FormulaFields.Item("IsLocalCurr").Text = 1
                Catch ex As Exception

                End Try
            End If

            Select Case cReportType
                Case AgeingType.ARAgeing
                    .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
                Case AgeingType.ARAgeingSummary
                    .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
                Case AgeingType.APAgeing
                    .RecordSelectionFormula = "{_NCM_AP_AGEING.USERNAME}='" & sAPAGERunningDate & "'"
                Case AgeingType.APAgeingSummary
                    .RecordSelectionFormula = "{_NCM_AP_AGEING.USERNAME}='" & sAPAGERunningDate & "'"
            End Select
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
            If (bIsExcel) Then
                Dim sPath As String = String.Empty
                sPath = sExcelFilePath
                rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
                Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
                objExcelOptions.ExcelUseConstantColumnWidth = False
                rpt.ExportOptions.FormatOptions = objExcelOptions
                Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                objOptions.DiskFileName = sPath
                rpt.ExportOptions.DestinationOptions = objOptions
                rpt.Export()
                System.Diagnostics.Process.Start(sPath)
            Else
                crViewer.ReportSource = rpt
                crViewer.Show()
            End If
        End If
 
    End Sub
    Friend Sub OPEN_HANADS_AGEING_5BUCKETS()
        Dim rpt As ReportDocument
        Dim subreport As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            Select Case ReportName
                Case ReportName.ARAging7B_Details
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing7B_HANADS
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary7B_HANADS
                    End Select

                Case ReportName.ARAging6B_Details
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing6B_HANADS
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary6B_HANADS
                    End Select
                Case Else
                    Select Case cReportType
                        Case AgeingType.ARAgeing
                            rpt = New rptARAgeing_HANADS
                        Case AgeingType.APAgeing
                            rpt = New rptAPAgeing_HANADS
                        Case AgeingType.ARAgeingSummary
                            rpt = New rptARAgeingSummary_HANADS
                        Case AgeingType.APAgeingSummary
                            rpt = New rptAPAgeingSummary_HANADS
                    End Select
            End Select
        End If

        With rpt

            Dim bIsUserDefinedRange As Boolean = False
            Dim iCount As Integer = 1
            Dim sFormat As String = "Bucket{0}Text"
            Dim sTemp As String = String.Format(sFormat, 1)

            Try
                .SetDataSource(cDataset)
                .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                bIsUserDefinedRange = True
            Catch ex As Exception
                bIsUserDefinedRange = False
            End Try

            If bIsUserDefinedRange Then
                Select Case ReportName
                    Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 6 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 7 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next

                    Case Else
                        sFormat = "Bucket{0}Text"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormula(sBucketText(iCount - 1))
                        Next

                        sFormat = "Bucket{0}Val"
                        For iCount = 1 To 5 Step 1
                            sTemp = String.Format(sFormat, iCount)
                            .DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                        Next
                End Select
            End If

            ' Sub Report (Only for AP Summary and AR Summary) --------------------------------------------
            If cReportType = AgeingType.APAgeingSummary Or cReportType = AgeingType.ARAgeingSummary Then
                subreport = .OpenSubreport("Summary")
                subreport.DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
                subreport.DataDefinition.FormulaFields.Item("LocalCurr").Text = SetStringIntoCrystalFormula(sLocalCurr)

                If bIsUserDefinedRange Then
                    sFormat = "Bucket{0}Val"
                    Select Case ReportName
                        Case ReportName.ARAging6B_Details, ReportName.ARAging6B_Summary
                            For iCount = 1 To 6 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case ReportName.ARAging7B_Details, ReportName.ARAging7B_Summary
                            For iCount = 1 To 7 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                        Case Else
                            For iCount = 1 To 5 Step 1
                                sTemp = String.Format(sFormat, iCount)
                                subreport.DataDefinition.FormulaFields.Item(sTemp).Text = sBucketVal(iCount - 1)
                            Next
                    End Select
                End If
            End If

            ' Sub Report Company Details -----------------------------------------------------------------
            Dim StatementDate As String = "DateValue(" & cAsAtDate.Substring(0, 4) & "," & cAsAtDate.Substring(4, 2) & "," & cAsAtDate.Substring(6, 2) & ")"
            subreport = .OpenSubreport("PageHeader.rpt")
            subreport.DataDefinition.FormulaFields.Item("CompanyId").Text = SetStringIntoCrystalFormula(oCompany.CompanyName)
            subreport.DataDefinition.FormulaFields.Item("BPCode").Text = SetStringIntoCrystalFormula(sBPCode)
            subreport.DataDefinition.FormulaFields.Item("BPCodeFr").Text = SetStringIntoCrystalFormula(sBPCodeFr)
            subreport.DataDefinition.FormulaFields.Item("BPCodeTo").Text = SetStringIntoCrystalFormula(sBPCodeTo)
            subreport.DataDefinition.FormulaFields.Item("BPGrpFr").Text = SetStringIntoCrystalFormula(sBPGrpFr)
            subreport.DataDefinition.FormulaFields.Item("BPGrpTo").Text = SetStringIntoCrystalFormula(sBPGrpTo)
            subreport.DataDefinition.FormulaFields.Item("SlsFr").Text = SetStringIntoCrystalFormula(sSlsFr)
            subreport.DataDefinition.FormulaFields.Item("SlsTo").Text = SetStringIntoCrystalFormula(sSlsTo)
            subreport.DataDefinition.FormulaFields.Item("AsAtDate").Text = StatementDate
            subreport.DataDefinition.FormulaFields.Item("AgingBy").Text = SetStringIntoCrystalFormula(sAgingBy)
            If (iSectionPageBreak = 0) Then
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("No")
            Else
                subreport.DataDefinition.FormulaFields.Item("PageBreak").Text = SetStringIntoCrystalFormula("Yes")
            End If
            ' Main Report Parameter ----------------------------------------------------------------------
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("AgeingBy").Text = cAgeBy
            .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
            .DataDefinition.FormulaFields.Item("PageBreak").Text = iSectionPageBreak

            If ReportName = ReportName.ARAging6B_Details Then
                .DataDefinition.FormulaFields.Item("IsLocalCurr").Text = 1
            End If

            Select Case cReportType
                Case AgeingType.APAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AP Ageing Details")
                Case AgeingType.APAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AP Ageing Summary")
                Case AgeingType.ARAgeing
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Details")
                Case AgeingType.ARAgeingSummary
                    subreport.DataDefinition.FormulaFields.Item("ReportTitle").Text = SetStringIntoCrystalFormula("AR Ageing Summary")
            End Select

            Select Case cReportType
                Case AgeingType.ARAgeing
                    .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
                Case AgeingType.ARAgeingSummary
                    .RecordSelectionFormula = "{_NCM_AR_AGEING.USERNAME}='" & sARAGERunningDate & "'"
                Case AgeingType.APAgeing
                    .RecordSelectionFormula = "{_NCM_AP_AGEING.USERNAME}='" & sAPAGERunningDate & "'"
                Case AgeingType.APAgeingSummary
                    .RecordSelectionFormula = "{_NCM_AP_AGEING.USERNAME}='" & sAPAGERunningDate & "'"
            End Select
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
            If (bIsExcel) Then
                Dim sPath As String = String.Empty
                sPath = sExcelFilePath
                rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
                Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
                objExcelOptions.ExcelUseConstantColumnWidth = False
                rpt.ExportOptions.FormatOptions = objExcelOptions
                Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                objOptions.DiskFileName = sPath
                rpt.ExportOptions.DestinationOptions = objOptions
                rpt.Export()
                System.Diagnostics.Process.Start(sPath)
            Else
                crViewer.ReportSource = rpt
                crViewer.Show()
            End If
        End If
    End Sub
    Friend Sub OPEN_HANADS_ARSOA()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            Select Case ReportName
                Case ReportName.ARSoa
                    If cIsLandscape = "Y" Then
                        rpt = New SA_LANDSCAPE_HANADS
                    Else
                        rpt = New SA_HANADS
                    End If
                Case ReportName.OMARSoa
                    rpt = New SA_OM_HANADS
            End Select
        End If

        ' Choose the correct ageing method
        Select Case cReport
            Case ReportCode.SOA_ByDocDate
                Me.Text = "Statement of Account - By Doc Date"
            Case ReportCode.SOA_ByDueDate
                Me.Text = "Statement of Account - By Due Date"
        End Select

        With rpt
            .SetDataSource(cDataset)
            subreport = .OpenSubreport("CH.rpt")
            If cHideLogo Then
                subreport.ReportDefinition.ReportObjects.Item("Field8").ObjectFormat.EnableSuppress = True
            Else
                subreport.ReportDefinition.ReportObjects.Item("Field8").ObjectFormat.EnableSuppress = False
            End If

            Dim StatementDate As String = "DateValue(" & AsAtDate.Substring(0, 4) & "," & AsAtDate.Substring(4, 2) & "," & AsAtDate.Substring(6, 2) & ")"
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("ReportType").Text = Report
            .DataDefinition.FormulaFields.Item("PeriodType").Text = Period
            .DataDefinition.FormulaFields.Item("IsBBF").Text = IsBBF
            .DataDefinition.FormulaFields.Item("IsSNP").Text = IsSNP
            .DataDefinition.FormulaFields.Item("IsGAT").Text = IsGAT
            .DataDefinition.FormulaFields.Item("IsHAS").Text = IsHAS
            .DataDefinition.FormulaFields.Item("IsHFN").Text = IsHFN

            If cHideHeader Then
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 1
            Else
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 0
            End If

            If sExportCardCode.Length > 0 Then
                .RecordSelectionFormula = "{_NCM_SOC.USERNAME} = '" & sARSOARunningDate & "' AND {_NCM_SOC.CARDCODE} = '" & sExportCardCode & "'"
            Else
                .RecordSelectionFormula = "{_NCM_SOC.USERNAME} = '" & sARSOARunningDate & "'"
            End If

            .Refresh()
        End With

        If bIsExport Then
            Dim sPath As String = String.Empty
            sPath = sExportPath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = oExportType
            If oExportType = ExportFormatType.Excel Then
                Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
                objExcelOptions.ExcelUseConstantColumnWidth = False
                rpt.ExportOptions.FormatOptions = objExcelOptions
            End If

            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub
    Friend Sub OPEN_HANADS_ARSOA_PROJ()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New SA_PROJ_HANADS
        End If

        ' Choose the correct ageing method
        Select Case cReport
            Case ReportCode.SOA_ByDocDate
                Me.Text = "Statement of Account (Group By Project) - By Doc Date"
            Case ReportCode.SOA_ByDueDate
                Me.Text = "Statement of Account (Group By Project) - By Due Date"
        End Select

        With rpt
            .SetDataSource(cDataset)
            subreport = .OpenSubreport("CH.rpt")
            If cHideLogo Then
                subreport.ReportDefinition.ReportObjects.Item("Field8").ObjectFormat.EnableSuppress = True
            Else
                subreport.ReportDefinition.ReportObjects.Item("Field8").ObjectFormat.EnableSuppress = False
            End If

            Dim StatementDate As String = "DateValue(" & AsAtDate.Substring(0, 4) & "," & AsAtDate.Substring(4, 2) & "," & AsAtDate.Substring(6, 2) & ")"
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("ReportType").Text = Report
            .DataDefinition.FormulaFields.Item("PeriodType").Text = Period
            .DataDefinition.FormulaFields.Item("IsBBF").Text = IsBBF
            .DataDefinition.FormulaFields.Item("IsSNP").Text = IsSNP
            .DataDefinition.FormulaFields.Item("IsGAT").Text = IsGAT
            .DataDefinition.FormulaFields.Item("IsHAS").Text = IsHAS
            .DataDefinition.FormulaFields.Item("IsHFN").Text = IsHFN

            If cHideHeader Then
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 1
            Else
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 0
            End If

            If sExportCardCode.Length > 0 Then
                .RecordSelectionFormula = "{_NCM_SOC.USERNAME} = '" & sARSOARunningDate & "' AND {_NCM_SOC.CARDCODE} = '" & sExportCardCode & "'"
            Else
                .RecordSelectionFormula = "{_NCM_SOC.USERNAME} = '" & sARSOARunningDate & "'"
            End If

            .Refresh()
        End With

        If bIsExport Then
            Dim sPath As String = String.Empty
            sPath = sExportPath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = oExportType
            If oExportType = ExportFormatType.Excel Then
                Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
                objExcelOptions.ExcelUseConstantColumnWidth = False
                rpt.ExportOptions.FormatOptions = objExcelOptions
            End If

            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub

    Friend Sub OPEN_HANADS_APSOA()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument
        Dim bIsIncludeLG_BankDetail As Boolean = False
        Dim StatementDate As String = "DateValue(" & AsAtDate.Substring(0, 4) & "," & AsAtDate.Substring(4, 2) & "," & AsAtDate.Substring(6, 2) & ")"

        ' Instantiate the correct report
        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New SA_AP_HANADS
        End If

        ' Choose the correct ageing method
        Select Case cReport
            Case ReportCode.SOA_ByDocDate
                Me.Text = "Statement of Account AP - By Doc Date"
            Case ReportCode.SOA_ByDueDate
                Me.Text = "Statement of Account AP - By Due Date"
        End Select

        ' Set all tables with correct login info
        With rpt
            .SetDataSource(cDataset)
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("ReportType").Text = Report
            .DataDefinition.FormulaFields.Item("PeriodType").Text = Period
            .DataDefinition.FormulaFields.Item("IsBBF").Text = IsBBF
            .DataDefinition.FormulaFields.Item("IsSNP").Text = IsSNP
            .DataDefinition.FormulaFields.Item("IsGAT").Text = IsGAT
            .DataDefinition.FormulaFields.Item("IsHAS").Text = IsHAS
            .DataDefinition.FormulaFields.Item("IsHFN").Text = IsHFN

            If cHideHeader Then
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 1
            Else
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 0
            End If

            .RecordSelectionFormula = "{_NCM_SOC_AP.USERNAME} = '" & sAPSOARunningDate & "'"
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
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If

     
    End Sub
    Friend Sub OPEN_HANADS_PAYMENTVOUCHER()
        Try
            Dim subreport As ReportDocument
            Dim rpt As ReportDocument

            Select Case cReportName
                Case ReportName.PV
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sReportName)
                    Else
                        rpt = New PV_HANADS
                    End If
                Case ReportName.RA
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sReportName)
                    Else
                        rpt = New RA_HANADS
                    End If
            End Select

            With rpt
                .SetDataSource(cDataset)
                subreport = .OpenSubreport("InvDtl")
                subreport.DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"

                If cShowDetails = True Then
                    subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'N'"
                Else
                    subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'Y'"
                End If

                .DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
                .DataDefinition.FormulaFields.Item("Username").Text = "'" & oCompany.UserName & "'"

                .RecordSelectionFormula = "{OVPM.DocNum} = " & cDocNum & " AND {OVPM.Series} = " & cSeries
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
                crViewer.ReportSource = rpt
                crViewer.DisplayGroupTree = False
                crViewer.PerformAutoScale()
                crViewer.Show()
            End If

        Catch ex As Exception
            MsgBox("[PrintPaymentVoucher] : " & ex.ToString)
        End Try
    End Sub
    Friend Sub OPEN_HANADS_DRAFTPV()
        Try
            Dim subreport As ReportDocument
            Dim rpt As ReportDocument

            If g_bIsShared Then
                rpt = New ReportDocument
                rpt.Load(g_sReportName)
            Else
                rpt = New DPV_HANADS
            End If

            With rpt
                .SetDataSource(cDataset)
                subreport = .OpenSubreport("InvDtl")
                If cShowDetails = True Then
                    subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'N'"
                Else
                    subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'Y'"
                End If
                subreport.DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
                subreport.RecordSelectionFormula = "{NCM_VIEW_DRAFTPV_INVOICE.PaymentDocEntry} = " & cDocEntry & " AND {NCM_VIEW_DRAFTPV_INVOICE.PaymentObjType} = '46' "

                ' =====================================================
                .DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
                .DataDefinition.FormulaFields.Item("Username").Text = "'" & oCompany.UserName & "'"
                .RecordSelectionFormula = "{OPDF.DocNum} = " & cDocNum & " AND {OPDF.Series} = " & cSeries & " AND {OPDF.DocEntry} = " & cDocEntry
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
                crViewer.ReportSource = rpt
                crViewer.DisplayGroupTree = False
                crViewer.PerformAutoScale()
                crViewer.Show()
            End If

        Catch ex As Exception
            MsgBox("DPV :" & ex.ToString)
        End Try
    End Sub
    Friend Sub OPEN_HANADS_PV_EMAIL()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sReportName)
        Else
            rpt = New PV_HANADS
        End If

        ' Choose the correct ageing method

        With rpt
            .SetDataSource(cDataset)
            subreport = .OpenSubreport("InvDtl")
            subreport.DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"

            If cShowDetails = True Then
                subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'N'"
            Else
                subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'Y'"
            End If

            .DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
            .DataDefinition.FormulaFields.Item("Username").Text = "'" & oCompany.UserName & "'"

            .RecordSelectionFormula = "{OVPM.DocNum} = " & cDocNum & " AND {OVPM.Series} = " & cSeries
            .Refresh()
        End With

        If bIsExport Then
            Dim sPath As String = String.Empty
            sPath = sExportPath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = oExportType
            If oExportType = ExportFormatType.Excel Then
                Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
                objExcelOptions.ExcelUseConstantColumnWidth = False
                rpt.ExportOptions.FormatOptions = objExcelOptions
            End If

            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub

    Friend Sub OPEN_HANADS_OFFICIALRECEIPT()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sReportName)
        Else
            rpt = New OR_HANADS
        End If

        With rpt
            .SetDataSource(cDataset)
            subreport = .OpenSubreport("InvDtl")

            If cShowDetails = True Then
                subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'N'"
            Else
                subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'Y'"
            End If

            subreport.DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
            .DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
            .DataDefinition.FormulaFields.Item("Username").Text = "'" & oCompany.UserName & "'"
            .RecordSelectionFormula = "{ORCT.DocNum} = " & cDocNum & " AND {ORCT.Series} = " & cSeries
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
            crViewer.ReportSource = rpt
            crViewer.DisplayGroupTree = False
            crViewer.PerformAutoScale()
            crViewer.Show()
        End If
 
    End Sub
    Friend Sub OPEN_HANADS_PV_RANGE()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sReportName)
        Else
            rpt = New RPV_HANADS
        End If

        With rpt
            .SetDataSource(cDataset)
            subreport = .OpenSubreport("InvDtl")
            subreport.DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"

            If cShowDetails = True Then
                subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'N'"
                .DataDefinition.FormulaFields.Item("SuppressDetail").Text = "'N'"
            Else
                subreport.DataDefinition.FormulaFields.Item("SuppressDtl").Text = "'Y'"
                .DataDefinition.FormulaFields.Item("SuppressDetail").Text = "'Y'"
            End If

            .DataDefinition.FormulaFields.Item("ShowTaxDate").Text = "'" & cShowTaxDate & "'"
            .DataDefinition.FormulaFields.Item("Username").Text = "'" & oCompany.UserName & "'"
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
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If

    End Sub
    Friend Sub OPEN_HANADS_BANKRECONCILIATION()
        Try
            Dim rpt As ReportDocument

            If g_bIsShared Then
                rpt = New ReportDocument
                rpt.Load(g_sSharedReportName)
            Else
                rpt = New RPT_BANK_RECON
            End If

            rpt.SetDataSource(cDataset)
            crViewer.Text = "Bank Reconciliation Report"
            crViewer.Name = "Bank Reconciliation Report"
            crViewer.DisplayGroupTree = False

            rpt.DataDefinition.FormulaFields.Item("InputAcct").Text = "'" & cBankAccount & "'"
            rpt.DataDefinition.FormulaFields.Item("InputDate").Text = "'" & cBankDate & "'"
            rpt.DataDefinition.FormulaFields.Item("InputYear").Text = "'" & cBankDate.Substring(0, 4) & "'"
            rpt.DataDefinition.FormulaFields.Item("InputMonth").Text = "'" & cBankDate.Substring(4, 2) & "'"
            rpt.DataDefinition.FormulaFields.Item("InputDay").Text = "'" & cBankDate.Substring(6, 2) & "'"

            Select Case cClientType
                Case "S"
                    ' Web Browser
                    ' generate pdf, put to a standard local folder, with specific name, call the pdf out.

                    rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                    rpt.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

                    Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                    objOptions.DiskFileName = g_sExportPath 'g_sExportPath
                    rpt.ExportOptions.DestinationOptions = objOptions
                    rpt.Export()

                Case "D"
                    crViewer.ReportSource = rpt
                    crViewer.Show()
            End Select
        Catch ex As Exception

        End Try
    End Sub

    Friend Sub OpenArSOA()
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        crConnectionInfo.ServerName = CONST_ODBC_SERVER_NAME
        crConnectionInfo.DatabaseName = oCompany.CompanyDB
        crConnectionInfo.UserID = global_DBUsername
        crConnectionInfo.Password = global_DBPassword

        Dim bIsIncludeLG_BankDetail As Boolean = False

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            Select Case ReportName
                Case ReportName.ARSoa
                    If cIsLandscape = "Y" Then
                        rpt = New SA_LANDSCAPE_HANA
                    Else
                        rpt = New SA_HANA
                    End If
                Case ReportName.OMARSoa
                    rpt = New SA_OM
            End Select
        End If

        ' Choose the correct ageing method
        Select Case cReport
            Case ReportCode.SOA_ByDocDate
                Me.Text = "Statement of Account - By Doc Date"
            Case ReportCode.SOA_ByDueDate
                Me.Text = "Statement of Account - By Due Date"
        End Select

        ' Set all tables with correct login info
        With rpt
            CrTables = .Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next
            subreport = .OpenSubreport("CH.rpt")
            CrTables = subreport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            If cHideLogo Then
                subreport.ReportDefinition.ReportObjects.Item("Field8").ObjectFormat.EnableSuppress = True
            Else
                subreport.ReportDefinition.ReportObjects.Item("Field8").ObjectFormat.EnableSuppress = False
            End If

            subreport = .OpenSubreport("Notes.rpt")
            CrTables = subreport.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            Next

            Dim StatementDate As String = "DateValue(" & AsAtDate.Substring(0, 4) & "," & AsAtDate.Substring(4, 2) & "," & AsAtDate.Substring(6, 2) & ")"
            .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
            .DataDefinition.FormulaFields.Item("ReportType").Text = Report
            .DataDefinition.FormulaFields.Item("PeriodType").Text = Period
            .DataDefinition.FormulaFields.Item("IsBBF").Text = IsBBF
            .DataDefinition.FormulaFields.Item("IsSNP").Text = IsSNP
            .DataDefinition.FormulaFields.Item("IsGAT").Text = IsGAT
            .DataDefinition.FormulaFields.Item("IsHAS").Text = IsHAS
            .DataDefinition.FormulaFields.Item("IsHFN").Text = IsHFN

            If (ReportName = ReportName.OMARSoa) Then
                .DataDefinition.FormulaFields.Item("LocalCurr").Text = "'" & sLocalCurr & "'"
                subreport = .OpenSubreport("Notes.rpt - 01")
                CrTables = subreport.Database.Tables
                For Each CrTable In CrTables
                    crtableLogoninfo = CrTable.LogOnInfo
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
                    CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                Next
            End If

            If cHideHeader Then
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 1
                '.ReportDefinition.ReportObjects.Item("Subreport1").ObjectFormat.EnableSuppress = True
            Else
                .DataDefinition.FormulaFields.Item("HideHeader").Text = 0
                '.ReportDefinition.ReportObjects.Item("Subreport1").ObjectFormat.EnableSuppress = False
            End If

            ''''Start LG Subreport
            subreport = New CrystalDecisions.CrystalReports.Engine.ReportDocument
            Try
                subreport = .OpenSubreport("SA_LG_BankDetail.rpt")
                bIsIncludeLG_BankDetail = True
            Catch ex As Exception
                bIsIncludeLG_BankDetail = False
            End Try
            If (subreport Is Nothing) Then
                bIsIncludeLG_BankDetail = False
            End If

            If (bIsIncludeLG_BankDetail) Then
                CrTables = subreport.Database.Tables
                For Each CrTable In CrTables
                    crtableLogoninfo = CrTable.LogOnInfo
                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
                    CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                Next
            End If
            ''''End LG Subreport

            If sExportCardCode.Length > 0 Then
                .RecordSelectionFormula = "{_NCM_SOC.USERNAME} = '" & sARSOARunningDate & "' AND {_NCM_SOC.CARDCODE} = '" & sExportCardCode & "'"
            Else
                .RecordSelectionFormula = "{_NCM_SOC.USERNAME} = '" & sARSOARunningDate & "'"
            End If
            .Refresh()
        End With

        If bIsExport Then
            Dim sPath As String = String.Empty
            sPath = sExportPath
            rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
            rpt.ExportOptions.ExportFormatType = oExportType
            If oExportType = ExportFormatType.Excel Then
                Dim objExcelOptions As New CrystalDecisions.Shared.ExcelFormatOptions
                objExcelOptions.ExcelUseConstantColumnWidth = False
                rpt.ExportOptions.FormatOptions = objExcelOptions
            End If

            Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
            objOptions.DiskFileName = sPath
            rpt.ExportOptions.DestinationOptions = objOptions
            rpt.Export()
        Else
            crViewer.ReportSource = rpt
            crViewer.Show()
        End If
    End Sub
    Friend Sub OPEN_HANADS_CHANGE_LOG()
        Dim subreport As ReportDocument
        Dim rpt As ReportDocument

        If g_bIsShared Then
            rpt = New ReportDocument
            rpt.Load(g_sSharedReportName)
        Else
            rpt = New rptChgLogAudit
        End If

        With rpt
            Me.Text = "Change Log Audit Report"
            .SetDataSource(cDataset)
            .DataDefinition.FormulaFields.Item("CompanyName").Text = "'" & CLOG_CompanyName.ToUpper & "'"
            .DataDefinition.FormulaFields.Item("UserID").Text = "'" & CLOG_UserId & "'"
            .DataDefinition.FormulaFields.Item("GenBy").Text = "'" & CLOG_GenBy & "'"

            If CLOG_dtFrom <> Nothing Then
                .DataDefinition.FormulaFields.Item("DtFrom").Text = String.Format("'{0}'", CLOG_dtFrom.ToShortDateString())
            End If

            If CLOG_dtTo <> Nothing Then
                .DataDefinition.FormulaFields.Item("DtTo").Text = String.Format("'{0}'", CLOG_dtTo.ToShortDateString())
            End If
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
            crViewer.ReportSource = rpt
            crViewer.DisplayGroupTree = False
            crViewer.PerformAutoScale()
            crViewer.Show()
        End If

    End Sub

    Private Sub frmViewer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim CrTables As Tables
            Dim CrTable As Table
            Dim subreport As ReportDocument
            Dim rpt As ReportDocument

            crConnectionInfo.ServerName = CONST_ODBC_SERVER_NAME
            crConnectionInfo.DatabaseName = oCompany.CompanyDB
            crConnectionInfo.UserID = global_DBUsername
            crConnectionInfo.Password = global_DBPassword

            Select Case ReportName
                Case ReportName.CHANGE_LOG_AUDIT
                    OPEN_HANADS_CHANGE_LOG()

                Case ReportName.ARAging_Details, ReportName.ARAging_Summary
                    OPEN_HANADS_AGEING_5BUCKETS()

                Case ReportName.APAging_Details, ReportName.APAging_Summary
                    OPEN_HANADS_AGEING_5BUCKETS()

                Case ReportName.PV
                    OPEN_HANADS_PAYMENTVOUCHER()

                Case ReportName.RA
                    OPEN_HANADS_PAYMENTVOUCHER()

                Case ReportName.PVDraft
                    OPEN_HANADS_DRAFTPV()

                Case ReportName.PV_Range
                    OPEN_HANADS_PV_RANGE()

                Case ReportName.IRA
                    OPEN_HANADS_OFFICIALRECEIPT()

                Case ReportName.ARAgeingDetailsCRM
                    OPEN_HANADS_AGEING_CRM()

                Case ReportName.ARAging_Details_Proj, ReportName.ARAging_Summary_Proj
                    OpenAgingReportProject_HANA()

                Case ReportName.APAging_Details_Proj, ReportName.APAging_Summary_Proj
                    OpenAgingReportProject_HANA()

                Case ReportName.ARAging6B_Details
                    OPEN_AGEING_REPORT_6_AND_7_BUCKETS()

                Case ReportName.ARAging7B_Details
                    OPEN_AGEING_REPORT_6_AND_7_BUCKETS()

                Case ReportName.ARSoa
                    OPEN_HANADS_ARSOA()

                Case ReportName.ARSOA_BY_PROJECT
                    OPEN_HANADS_ARSOA_PROJ()

                Case ReportName.APSoa
                    OPEN_HANADS_APSOA()

                Case ReportName.OMARSoa
                    OpenArSOA()

                Case ReportName.MRPSupplyDemandReport

                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RPT_MRP
                    End If
                    rpt.SetDataSource(cDataset)
                    crViewer.Text = "MRP Demand and Supply Report"
                    crViewer.Name = "MRP Demand and Supply Report"
                    crViewer.ReportSource = rpt
                    crViewer.Show()

                Case ReportName.Items_ABC_Analysis
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RPT_ABC
                    End If

                    rpt.SetDataSource(cDataset)
                    crViewer.Text = "Items ABC Analysis Report"
                    crViewer.Name = "Items ABC Analysis Report"
                    crViewer.ReportSource = rpt
                    crViewer.Show()

                Case ReportName.ReOrder_Level_Recommendation
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RPT_RLR
                    End If
                    rpt.SetDataSource(cDataset)
                    rpt.DataDefinition.FormulaFields.Item("SmoothFactor").Text = "'" & Math.Round(dSmoothFactor, 2) & "'"
                    rpt.DataDefinition.FormulaFields.Item("NoOfWeek").Text = "'" & iNoOfWeek & "'"
                    crViewer.Name = "ReOrder Level Recommendation Report"
                    crViewer.ReportSource = rpt
                    crViewer.Show()

                Case ReportName.Weighted_Average_Demand
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RPT_WAR
                    End If
                    rpt.SetDataSource(cDataset)
                    crViewer.Name = "Weighted Average Demand Report"
                    crViewer.ReportSource = rpt
                    crViewer.Show()

                Case ReportName.SOA_TOS
                    ' ---------------------------------------------------------------------------------------------------------**
                    ' Instantiate the correct report
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New SA_TOS_HANA
                    End If

                    ' Choose the correct ageing method
                    Select Case cReport
                        Case ReportCode.SOA_ByDocDate
                            Me.Text = "Statement of Account (Project) - By Doc Date"
                        Case ReportCode.SOA_ByDueDate
                            Me.Text = "Statement of Account (Project) - By Due Date"
                    End Select

                    ' Set all tables with correct login info
                    With rpt
                        CrTables = .Database.Tables
                        For Each CrTable In CrTables
                            crtableLogoninfo = CrTable.LogOnInfo
                            crtableLogoninfo.ConnectionInfo = crConnectionInfo
                            CrTable.ApplyLogOnInfo(crtableLogoninfo)
                            CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                        Next
                        subreport = .OpenSubreport("CH.rpt")
                        CrTables = subreport.Database.Tables
                        For Each CrTable In CrTables
                            crtableLogoninfo = CrTable.LogOnInfo
                            crtableLogoninfo.ConnectionInfo = crConnectionInfo
                            CrTable.ApplyLogOnInfo(crtableLogoninfo)
                            CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                        Next
                        subreport = .OpenSubreport("Notes.rpt")
                        CrTables = subreport.Database.Tables
                        For Each CrTable In CrTables
                            crtableLogoninfo = CrTable.LogOnInfo
                            crtableLogoninfo.ConnectionInfo = crConnectionInfo
                            CrTable.ApplyLogOnInfo(crtableLogoninfo)
                            CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                        Next

                        Dim StatementDate As String = "DateValue(" & AsAtDate.Substring(0, 4) & "," & AsAtDate.Substring(4, 2) & "," & AsAtDate.Substring(6, 2) & ")"
                        .DataDefinition.FormulaFields.Item("statementDate").Text = StatementDate
                        .DataDefinition.FormulaFields.Item("ReportType").Text = Report
                        .DataDefinition.FormulaFields.Item("PeriodType").Text = Period
                        .DataDefinition.FormulaFields.Item("IsBBF").Text = IsBBF
                        .DataDefinition.FormulaFields.Item("IsSNP").Text = IsSNP
                        .DataDefinition.FormulaFields.Item("IsGAT").Text = IsGAT
                        .DataDefinition.FormulaFields.Item("IsHAS").Text = IsHAS
                        .DataDefinition.FormulaFields.Item("IsHFN").Text = IsHFN
                        .DataDefinition.FormulaFields.Item("UserName").Text = "'" & oCompany.UserName & "'"
                        .DataDefinition.FormulaFields.Item("Business").Text = "'" & cBusiness.ToUpper & "'"

                        .Refresh()
                    End With
                    crViewer.ReportSource = rpt
                    crViewer.Show()
                    ' ---------------------------------------------------------------------------------------------------------**

                Case ReportName.GPA
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New rptGPA
                    End If
                    With rpt
                        Me.Text = "Gross Profit Analysis"
                        .DataDefinition.FormulaFields.Item("CompanyName").Text = "'" & oCompany.CompanyName & "'"
                        .DataDefinition.FormulaFields.Item("UserId").Text = "'" & oCompany.UserName & "'"
                        .DataDefinition.FormulaFields.Item("Title").Text = "'Gross Profit Analysis'"
                        .DataDefinition.FormulaFields.Item("FromDate").Text = "'" & cFrDate & "'"
                        .DataDefinition.FormulaFields.Item("ToDate").Text = "'" & cToDate & "'"
                        .DataDefinition.FormulaFields.Item("FromBP").Text = "'" & cFrBP & "'"
                        .DataDefinition.FormulaFields.Item("ToBP").Text = "'" & cToBP & "'"
                        .DataDefinition.FormulaFields.Item("FromProj").Text = "'" & cFrProj & "'"
                        .DataDefinition.FormulaFields.Item("ToProj").Text = "'" & cToProj & "'"

                        rpt.SetDataSource(cDataset)
                        crViewer.ReportSource = rpt
                    End With
                    SBO_Application.StatusBar.SetText("Operation ended successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    ' ---------------------------------------------------------------------------------------------------------**
                Case ReportName.ARPayment
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RptPayment
                    End If
                    If Not (cReportDataSet Is Nothing) Then
                        rpt.SetDataSource(cReportDataSet)
                    End If
                    rpt.DataDefinition.FormulaFields.Item("CompanyName").Text = "'" & oCompany.CompanyName.ToUpper & "'"
                    rpt.DataDefinition.FormulaFields.Item("DocDateFrom").Text = "'" & cDateFrom & "'"
                    rpt.DataDefinition.FormulaFields.Item("DocDateTo").Text = "'" & cDateTo & "'"
                    rpt.DataDefinition.FormulaFields.Item("CardCodeFrom").Text = "'" & cBPFrom & "'"
                    rpt.DataDefinition.FormulaFields.Item("CardCodeTo").Text = "'" & cBPTo & "'"
                    rpt.DataDefinition.FormulaFields.Item("ReportTitle").Text = "'INCOMING PAYMENT REPORT'"
                    rpt.DataDefinition.FormulaFields.Item("ParamBPLabel").Text = "'Customer'"
                    crViewer.ReportSource = rpt

                Case ReportName.APPayment
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RptPayment
                    End If
                    If Not (cReportDataSet Is Nothing) Then
                        rpt.SetDataSource(cReportDataSet)
                    End If
                    rpt.DataDefinition.FormulaFields.Item("CompanyName").Text = "'" & oCompany.CompanyName.ToUpper & "'"
                    rpt.DataDefinition.FormulaFields.Item("DocDateFrom").Text = "'" & cDateFrom & "'"
                    rpt.DataDefinition.FormulaFields.Item("DocDateTo").Text = "'" & cDateTo & "'"
                    rpt.DataDefinition.FormulaFields.Item("CardCodeFrom").Text = "'" & cBPFrom & "'"
                    rpt.DataDefinition.FormulaFields.Item("CardCodeTo").Text = "'" & cBPTo & "'"
                    rpt.DataDefinition.FormulaFields.Item("ReportTitle").Text = "'OUTGOING PAYMENT REPORT'"
                    rpt.DataDefinition.FormulaFields.Item("ParamBPLabel").Text = "'Vendor'"
                    crViewer.ReportSource = rpt

                Case Inecom_SDK_Reporting_Package.ReportName.BankReconciliation
                    OPEN_HANADS_BANKRECONCILIATION()

                Case ReportName.SO_Detail_Proj
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RPT_SO_PROJ
                    End If
                    rpt.SetDataSource(cDataset)
                    crViewer.Text = "SO Project Detail By Customer"
                    crViewer.Name = "SO Project Detail By Customer"
                    crViewer.ReportSource = rpt
                    crViewer.Show()

                Case ReportName.PO_Detail_Proj
                    If g_bIsShared Then
                        rpt = New ReportDocument
                        rpt.Load(g_sSharedReportName)
                    Else
                        rpt = New RPT_PO_PROJ
                    End If
                    rpt.SetDataSource(cDataset)
                    crViewer.Text = "PO Project Detail By Vendor"
                    crViewer.Name = "PO Project Detail By Vendor"
                    crViewer.ReportSource = rpt
                    crViewer.Show()

                Case Else
                    rpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                    crViewer.ReportSource = rpt
                    crViewer.Show()
            End Select
        Catch ex As Exception
            MessageBox.Show("[frmViewer].[frmViewer_Load]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function SetStringIntoCrystalFormula(ByVal InputString As String) As String
        If (InputString Is Nothing) Then
            Return String.Empty
        Else
            Return "'" & InputString.Replace("'", "''") & "'"
        End If
    End Function
End Class
