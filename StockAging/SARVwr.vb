Option Strict Off
Option Explicit On 

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class SARVwr
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
    Friend WithEvents crViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents crViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.crViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.SuspendLayout()
        '
        'crViewer
        '
        Me.crViewer.ActiveViewIndex = -1
        Me.crViewer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crViewer.Location = New System.Drawing.Point(0, 0)
        Me.crViewer.Name = "crViewer"
        Me.crViewer.ReportSource = Nothing
        Me.crViewer.Size = New System.Drawing.Size(480, 437)
        Me.crViewer.TabIndex = 0
        '
        'SARVwr
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 437)
        Me.Controls.Add(Me.crViewer)
        Me.Name = "SARVwr"
        Me.Text = "Report Viewer"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cServerName As String
    Private cDatabase As String
    Private cDBUser As String
    Private cDBPassword As String
    Private cUserName As String
    Private sItemFr, sItemTo As String
    Private sWareFr, sWareTo, sExclude, sAsAtDate As String
    Private g_sSharedReportName As String
    Private bIsShared As Boolean = False
    Private bIsExcel As Boolean = False
    Private sReportType As String = String.Empty
    Private sGroupBy As String = String.Empty
    Private sItemGrpFr As String = String.Empty
    Private sItemGrpTo As String = String.Empty
    Private sExcelFilePath As String = String.Empty
    Private sBucketText As String() = New String(10) {}
    Private iReportName As Integer
    Private sAgeType As String = "-1"
    Private dDataset As Dataset
    Private g_sExportPath As String = ""
    Private cClientType As String = ""

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
    Friend Property Dataset() As Dataset
        Get
            Return dDataset
        End Get
        Set(ByVal Value As Dataset)
            dDataset = Value
        End Set
    End Property
    Friend Property ItemFrom() As String
        Get
            Return sItemFr
        End Get
        Set(ByVal Value As String)
            sItemFr = Value
        End Set
    End Property
    Friend Property ItemTo() As String
        Get
            Return sItemTo
        End Get
        Set(ByVal Value As String)
            sItemTo = Value
        End Set
    End Property
    Friend Property WarehouseFrom() As String
        Get
            Return sWareFr
        End Get
        Set(ByVal Value As String)
            sWareFr = Value
        End Set
    End Property
    Friend Property WarehouseTo() As String
        Get
            Return sWareTo
        End Get
        Set(ByVal Value As String)
            sWareTo = Value
        End Set
    End Property
    Friend Property ExcludeZeroBalance() As String
        Get
            Return sExclude
        End Get
        Set(ByVal Value As String)
            sExclude = Value
        End Set
    End Property
    Friend Property ReportName() As Integer
        Get
            Return iReportName
        End Get
        Set(ByVal Value As Integer)
            iReportName = Value
        End Set
    End Property
    Friend Property SharedReportName() As String
        Get
            Return g_sSharedReportName
        End Get
        Set(ByVal Value As String)
            g_sSharedReportName = Value
        End Set
    End Property
    Friend Property AsAtDate() As String
        Get
            Return sAsAtDate
        End Get
        Set(ByVal Value As String)
            sAsAtDate = Value
        End Set
    End Property
    Friend Property IsShared() As Boolean
        Get
            Return bIsShared
        End Get
        Set(ByVal Value As Boolean)
            bIsShared = Value
        End Set
    End Property
    Friend Property GroupBy() As String
        Get
            Return sGroupBy
        End Get
        Set(ByVal Value As String)
            sGroupBy = Value
        End Set
    End Property
    Friend Property ReportType() As String
        Get
            Return sReportType
        End Get
        Set(ByVal Value As String)
            sReportType = Value
        End Set
    End Property
    Friend Property ItemGroupFrom() As String
        Get
            Return sItemGrpFr
        End Get
        Set(ByVal Value As String)
            sItemGrpFr = Value
        End Set
    End Property
    Friend Property ItemGroupTo() As String
        Get
            Return sItemGrpTo
        End Get
        Set(ByVal Value As String)
            sItemGrpTo = Value
        End Set
    End Property
    Friend Property IsExcel() As Boolean
        Get
            Return bIsExcel
        End Get
        Set(ByVal Value As Boolean)
            bIsExcel = Value
        End Set
    End Property
    Friend Property AgeType() As String
        Get
            Return sAgeType
        End Get
        Set(ByVal Value As String)
            sAgeType = Value
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

    Friend Sub OpenSummaryReport_FIFO()
        Try
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim SelectionFormula As String = String.Empty
            Dim CrTables As Tables
            Dim CrTable As Table
            Dim rpt As ReportDocument
            Dim CrField1, CrField2 As CrystalDecisions.CrystalReports.Engine.DatabaseFieldDefinition
            Dim CrGroups As CrystalDecisions.CrystalReports.Engine.Groups
            Dim CrGroup As CrystalDecisions.CrystalReports.Engine.Group
            Dim sDebugString As String = String.Empty

            crConnectionInfo.ServerName = cServerName
            crConnectionInfo.DatabaseName = cDatabase
            crConnectionInfo.UserID = global_DBUsername
            crConnectionInfo.Password = global_DBPassword

            If (bIsShared) Then
                rpt = New ReportDocument
                rpt.Load(g_sSharedReportName)
            Else
                rpt = New FIFO_COM_SUMM_HANA
            End If

            With rpt
                .DataDefinition.FormulaFields.Item("ItemFr").Text = "'" & sItemFr & "'"
                .DataDefinition.FormulaFields.Item("ItemTo").Text = "'" & sItemTo & "'"
                .DataDefinition.FormulaFields.Item("WHFr").Text = "'" & sWareFr & "'"
                .DataDefinition.FormulaFields.Item("WHTo").Text = "'" & sWareTo & "'"
                .DataDefinition.FormulaFields.Item("ZeroBalance").Text = "'" & sExclude & "'"
                .DataDefinition.FormulaFields.Item("DateString").Text = "'" & sAsAtDate & "'"
                .DataDefinition.FormulaFields.Item("ReportType").Text = "'" & sReportType & "'"
                .DataDefinition.FormulaFields.Item("ItemGrpFrom").Text = "'" & sItemGrpFr & "'"
                .DataDefinition.FormulaFields.Item("ItemGrpTo").Text = "'" & sItemGrpTo & "'"
                .DataDefinition.FormulaFields.Item("IsExcel").Text = IIf(bIsExcel, 1, 0)

                CrField1 = .Database.Tables("DS_FIFO").Fields("ItemCode")
                CrField2 = .Database.Tables("DS_FIFO").Fields("WhsCode")
                CrGroups = .DataDefinition.Groups

                If (String.Compare(sGroupBy, "ItemCode", True) = 0) Then
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField1
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField2
                        End If
                    Next
                Else

                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField2
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField1
                        End If
                    Next
                End If

                CrTables = .Database.Tables
                For Each CrTable In CrTables
                    If CrTable.Name <> "DS_FIFO" Then
                        crtableLogoninfo = CrTable.LogOnInfo
                        crtableLogoninfo.ConnectionInfo = crConnectionInfo
                        CrTable.ApplyLogOnInfo(crtableLogoninfo)
                        sDebugString = "[Main] " & oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                        CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                    End If
                Next
                .SetDataSource(dDataset)
                .Refresh()
            End With

            If (bIsExcel) Then
                Dim sPath As String = String.Empty
                sPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\Inecom Report\"
                If (Not System.IO.Directory.Exists(sPath)) Then
                    System.IO.Directory.CreateDirectory(sPath)
                End If
                sPath = sPath & "Stock Aging\"
                If (Not System.IO.Directory.Exists(sPath)) Then
                    System.IO.Directory.CreateDirectory(sPath)
                End If
                sPath = sPath & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xls"
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
                    crViewer.PerformAutoScale()
                    crViewer.Show()
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub OpenDetailReport_FIFO()
        Try
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim SelectionFormula As String = String.Empty
            Dim CrTables As Tables
            Dim CrTable As Table
            Dim rpt As ReportDocument
            Dim CrField1, CrField2 As CrystalDecisions.CrystalReports.Engine.DatabaseFieldDefinition
            Dim CrGroups As CrystalDecisions.CrystalReports.Engine.Groups
            Dim CrGroup As CrystalDecisions.CrystalReports.Engine.Group
            Dim sDebugString As String = String.Empty

            crConnectionInfo.ServerName = cServerName
            crConnectionInfo.DatabaseName = cDatabase
            crConnectionInfo.UserID = global_DBUsername
            crConnectionInfo.Password = global_DBPassword

            If (bIsShared) Then
                rpt = New ReportDocument
                rpt.Load(g_sSharedReportName)
            Else
                rpt = New FIFO_COM_DET_HANA
            End If

            With rpt
                .DataDefinition.FormulaFields.Item("ItemFr").Text = "'" & sItemFr & "'"
                .DataDefinition.FormulaFields.Item("ItemTo").Text = "'" & sItemTo & "'"
                .DataDefinition.FormulaFields.Item("WHFr").Text = "'" & sWareFr & "'"
                .DataDefinition.FormulaFields.Item("WHTo").Text = "'" & sWareTo & "'"
                .DataDefinition.FormulaFields.Item("ZeroBalance").Text = "'" & sExclude & "'"
                .DataDefinition.FormulaFields.Item("DateString").Text = "'" & sAsAtDate & "'"
                .DataDefinition.FormulaFields.Item("ReportType").Text = "'" & sReportType & "'"
                .DataDefinition.FormulaFields.Item("IsExcel").Text = IIf(bIsExcel, 1, 0)

                CrField1 = .Database.Tables("DS_FIFO").Fields("ItemCode")
                CrField2 = .Database.Tables("DS_FIFO").Fields("WhsCode")
                CrGroups = .DataDefinition.Groups

                If (String.Compare(sGroupBy, "ItemCode", True) = 0) Then
                    'Added: V03.02.2005
                    .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMCODE'"
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField1
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField2
                        End If
                    Next
                Else
                    'Added: V03.02.2005
                    .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMCODE'"
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField2
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField1
                        End If
                    Next
                End If

                CrTables = .Database.Tables
                For Each CrTable In CrTables
                    If CrTable.Name <> "DS_FIFO" Then
                        crtableLogoninfo = CrTable.LogOnInfo
                        crtableLogoninfo.ConnectionInfo = crConnectionInfo
                        CrTable.ApplyLogOnInfo(crtableLogoninfo)
                        sDebugString = "[Main] " & oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                        CrTable.Location = oCompany.CompanyDB & "." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                    End If
                Next
                .Refresh()
                .SetDataSource(dDataset)
            End With

            If (bIsExcel) Then
                Dim sPath As String = String.Empty
                sPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\Inecom Report\"
                If (Not System.IO.Directory.Exists(sPath)) Then
                    System.IO.Directory.CreateDirectory(sPath)
                End If
                sPath = sPath & "Stock Aging\"
                If (Not System.IO.Directory.Exists(sPath)) Then
                    System.IO.Directory.CreateDirectory(sPath)
                End If
                sPath = sPath & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xls"
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
                    crViewer.PerformAutoScale()
                    crViewer.Show()
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function SetStringIntoCrystalFormula(ByVal InputString As String) As String
        If (InputString Is Nothing) Then
            Return String.Empty
        Else
            Return "'" & InputString.Replace("'", "''") & "'"
        End If
    End Function
    Private Function SetStringIntoCrystalFormulaB(ByVal InputString As String) As String
        If (InputString Is Nothing) Then
            Return String.Empty
        Else
            Dim lastIndex As Integer = -1
            Dim outputString As String = String.Empty
            InputString = InputString.Trim

            lastIndex = InputString.LastIndexOf(" ")


            'Dim inString As String = String.Empty
            'Dim outString As String = String.Empty
            'Dim lastIndex As Integer = -1
            'inString = TextBox1.Text
            'lastIndex = inString.LastIndexOf(" ")
            'outString = inString.Substring(0, lastIndex) & " + chr(10) + " & inString.Substring(lastIndex + 1, inString.Length - 1 - lastIndex)
            'MsgBox(outString)

            If (lastIndex <> -1) Then
                outputString = "'" & InputString.Substring(0, lastIndex).TrimEnd().Replace("'", "''") & "' + chr(10) + '" & InputString.Substring(lastIndex + 1, InputString.Length - 1 - lastIndex).TrimStart().Replace("'", "''") & "'"
                Return outputString
            Else
                Return "'" & InputString.Replace("'", "''") & "'"
            End If
        End If
    End Function

    Friend Sub OpenSummaryReport_MOV()
        Try
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim SelectionFormula As String = String.Empty
            Dim rpt As ReportDocument
            Dim CrField1, CrField2, CrField3 As CrystalDecisions.CrystalReports.Engine.DatabaseFieldDefinition
            Dim CrGroups As CrystalDecisions.CrystalReports.Engine.Groups
            Dim CrGroup As CrystalDecisions.CrystalReports.Engine.Group
            Dim sDebugString As String = String.Empty

            If (bIsShared) Then
                rpt = New ReportDocument
                rpt.Load(g_sSharedReportName)
            Else
                rpt = New FIFO_COM_SUMM_HANADS
            End If
            With rpt
                .SetDataSource(dDataset)
                .DataDefinition.FormulaFields.Item("ItemFr").Text = "'" & sItemFr & "'"
                .DataDefinition.FormulaFields.Item("ItemTo").Text = "'" & sItemTo & "'"
                .DataDefinition.FormulaFields.Item("WHFr").Text = "'" & sWareFr & "'"
                .DataDefinition.FormulaFields.Item("WHTo").Text = "'" & sWareTo & "'"
                .DataDefinition.FormulaFields.Item("ZeroBalance").Text = "'" & sExclude & "'"
                .DataDefinition.FormulaFields.Item("DateString").Text = "'" & sAsAtDate & "'"
                .DataDefinition.FormulaFields.Item("ReportType").Text = "'" & sReportType & "'"
                .DataDefinition.FormulaFields.Item("ItemGrpFrom").Text = "'" & sItemGrpFr & "'"
                .DataDefinition.FormulaFields.Item("ItemGrpTo").Text = "'" & sItemGrpTo & "'"
                .DataDefinition.FormulaFields.Item("IsExcel").Text = IIf(bIsExcel, 1, 0)
                .DataDefinition.FormulaFields.Item("AgeType").Text = sAgeType

                'Added: V03.05.2005
                Dim iCount As Integer = 1
                Dim sFormat As String = "Bucket{0}Text"
                Dim sTemp As String = String.Empty

                For iCount = 1 To 9
                    sTemp = String.Format(sFormat, iCount)
                    .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormulaB(sBucketText(iCount - 1))
                Next

                CrField1 = .Database.Tables("DS_FIFO").Fields("ItemCode")
                CrField2 = .Database.Tables("DS_FIFO").Fields("WhsCode")
                CrField3 = .Database.Tables("DS_FIFO").Fields("ItemGroup")
                CrGroups = .DataDefinition.Groups

                If (String.Compare(sGroupBy, "ItemCode", True) = 0) Then
                    .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMCODE'"
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField1
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField1
                        End If
                    Next
                ElseIf (String.Compare(sGroupBy, "WhsCode", True) = 0) Then
                    .DataDefinition.FormulaFields.Item("GroupBy").Text = "'WHSCODE'"
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField2
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField1
                        End If
                    Next
                ElseIf (String.Compare(sGroupBy, "ItemCode_Whse", True) = 0) Then
                    .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMWHSE'"
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField1
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField2
                        End If
                    Next
                ElseIf (String.Compare(sGroupBy, "ItemGroup", True) = 0) Then
                    .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMGROUP'"
                    For Each CrGroup In CrGroups
                        If (CrGroup.ConditionField.Name = "Group1") Then
                            CrGroup.ConditionField = CrField3
                        End If
                        If (CrGroup.ConditionField.Name = "Group2") Then
                            CrGroup.ConditionField = CrField3
                        End If
                    Next
                End If

                .Refresh()
            End With


            If (bIsExcel) Then
                Dim sPath As String = String.Empty
                sPath = sExcelFilePath
                rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
                Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                objOptions.DiskFileName = sPath
                rpt.ExportOptions.DestinationOptions = objOptions
                rpt.Export()
                System.Diagnostics.Process.Start(sPath)
            Else
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
                    crViewer.PerformAutoScale()
                    crViewer.Show()
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub OpenDetailReport_MOV()
        Try
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim SelectionFormula As String = String.Empty
            Dim rpt As ReportDocument
            Dim CrField1, CrField2, CrField3 As CrystalDecisions.CrystalReports.Engine.DatabaseFieldDefinition
            Dim CrGroups As CrystalDecisions.CrystalReports.Engine.Groups
            Dim CrGroup As CrystalDecisions.CrystalReports.Engine.Group
            Dim sDebugString As String = String.Empty

            If (bIsShared) Then
                rpt = New ReportDocument
                rpt.Load(g_sSharedReportName)
            Else
                rpt = New FIFO_COM_DET_HANADS
            End If

            With rpt
                .SetDataSource(dDataset)
                .DataDefinition.FormulaFields.Item("ItemFr").Text = "'" & sItemFr & "'"
                .DataDefinition.FormulaFields.Item("ItemTo").Text = "'" & sItemTo & "'"
                .DataDefinition.FormulaFields.Item("WHFr").Text = "'" & sWareFr & "'"
                .DataDefinition.FormulaFields.Item("WHTo").Text = "'" & sWareTo & "'"
                .DataDefinition.FormulaFields.Item("ZeroBalance").Text = "'" & sExclude & "'"
                .DataDefinition.FormulaFields.Item("DateString").Text = "'" & sAsAtDate & "'"
                .DataDefinition.FormulaFields.Item("ReportType").Text = "'" & sReportType & "'"
                .DataDefinition.FormulaFields.Item("IsExcel").Text = IIf(bIsExcel, 1, 0)
                .DataDefinition.FormulaFields.Item("AgeType").Text = sAgeType

                Dim iCount As Integer = 1
                Dim sFormat As String = "Bucket{0}Text"
                Dim sTemp As String = String.Empty
                For iCount = 1 To 9
                    sTemp = String.Format(sFormat, iCount)
                    .DataDefinition.FormulaFields.Item(sTemp).Text = SetStringIntoCrystalFormulaB(sBucketText(iCount - 1))
                Next

                CrField1 = .Database.Tables("DS_FIFO").Fields("ItemCode")
                CrField2 = .Database.Tables("DS_FIFO").Fields("WhsCode")
                CrField3 = .Database.Tables("DS_FIFO").Fields("ItemGroup")
                CrGroups = .DataDefinition.Groups

                Select Case sGroupBy
                    Case "ItemCode"
                        'Added: V03.02.2005
                        .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMCODE'"
                        For Each CrGroup In CrGroups
                            If (CrGroup.ConditionField.Name = "Group1") Then
                                CrGroup.ConditionField = CrField1
                            End If
                            If (CrGroup.ConditionField.Name = "Group2") Then
                                CrGroup.ConditionField = CrField1
                            End If
                        Next
                    Case "WhsCode"
                        .DataDefinition.FormulaFields.Item("GroupBy").Text = "'WHSCODE'"
                        For Each CrGroup In CrGroups
                            If (CrGroup.ConditionField.Name = "Group1") Then
                                CrGroup.ConditionField = CrField2
                            End If
                            If (CrGroup.ConditionField.Name = "Group2") Then
                                CrGroup.ConditionField = CrField1
                            End If
                        Next
                    Case "ItemCode_Whse"
                        .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMWHSE'"
                        For Each CrGroup In CrGroups
                            If (CrGroup.ConditionField.Name = "Group1") Then
                                CrGroup.ConditionField = CrField1
                            End If
                            If (CrGroup.ConditionField.Name = "Group2") Then
                                CrGroup.ConditionField = CrField2
                            End If
                        Next
                    Case "ItemGroup"
                        .DataDefinition.FormulaFields.Item("GroupBy").Text = "'ITEMGROUP'"
                        For Each CrGroup In CrGroups
                            If (CrGroup.ConditionField.Name = "Group1") Then
                                CrGroup.ConditionField = CrField3
                            End If
                            If (CrGroup.ConditionField.Name = "Group2") Then
                                CrGroup.ConditionField = CrField3
                            End If
                        Next
                End Select

                .Refresh()

            End With

            If (bIsExcel) Then
                Dim sPath As String = String.Empty
                sPath = sExcelFilePath
                rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                rpt.ExportOptions.ExportFormatType = ExportFormatType.Excel
                Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                objOptions.DiskFileName = sPath
                rpt.ExportOptions.DestinationOptions = objOptions
                rpt.Export()
                System.Diagnostics.Process.Start(sPath)
            Else

                Select Case cClientType
                    Case "S"
                        ' Web Browser
                        ' generate pdf, put to a standard local folder, with specific name.
                        ' call the pdf out.

                        rpt.ExportOptions.ExportDestinationType = ExportDestinationType.DiskFile
                        rpt.ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat

                        Dim objOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                        objOptions.DiskFileName = g_sExportPath
                        rpt.ExportOptions.DestinationOptions = objOptions
                        rpt.Export()
                    Case Else
                        'Desktop
                        crViewer.ReportSource = rpt
                        crViewer.PerformAutoScale()
                        crViewer.Show()
                End Select

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub frmViewer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sDebugString As String = String.Empty
        Try
            'VARIABLE FOR CONNECTION PROPERTIES
            Dim crtableLogoninfos As New TableLogOnInfos
            Dim crtableLogoninfo As New TableLogOnInfo
            Dim crConnectionInfo As New ConnectionInfo
            Dim SelectionFormula As String = String.Empty

            ' crConnectionInfo.ServerName = oCompany.Server
            crConnectionInfo.ServerName = "TEST_HANA_USER"
            crConnectionInfo.DatabaseName = oCompany.CompanyDB
            crConnectionInfo.UserID = global_DBUsername
            crConnectionInfo.Password = global_DBPassword

            Select Case iReportName
                Case 9 'FIFO Non-Batch Items - SUMM
                    OpenSummaryReport_FIFO()
                Case 10 'Stock AGing
                    OpenDetailReport_FIFO()
                Case 11
                    OpenSummaryReport_MOV()
                Case 12
                    OpenDetailReport_MOV()
            End Select
        Catch ex As Exception
            MsgBox(sDebugString & vbCrLf & ex.ToString)
        End Try
    End Sub
End Class
