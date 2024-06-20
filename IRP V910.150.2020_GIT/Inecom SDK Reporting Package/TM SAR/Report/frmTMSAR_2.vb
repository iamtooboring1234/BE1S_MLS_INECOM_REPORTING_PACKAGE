Public Class frmTMSAR_2
    Dim _ds As DataSet = Nothing
    Dim _isRptShared As Boolean = False
    Dim _sRptPath As String = String.Empty

    Dim _AsAtDate As DateTime = DateTime.Now
    Dim _sItemFr As String = String.Empty
    Dim _sItemTo As String = String.Empty
    Dim _sItemGrpFr As String = String.Empty
    Dim _sItemGrpTo As String = String.Empty
    Dim _sWhseFr As String = String.Empty
    Dim _sWhseTo As String = String.Empty
    Dim _sDim1Fr As String = String.Empty
    Dim _sDim1To As String = String.Empty
    Dim _sDim2Fr As String = String.Empty
    Dim _sDim2To As String = String.Empty
    Dim _sItemTypeFr As String = String.Empty
    Dim _sItemTypeTo As String = String.Empty
    Dim _LocalCurr As String = String.Empty
    Dim _sReportType As String = String.Empty

#Region "Properties"
    Public Property crDateSet() As DataSet
        Get
            Return _ds
        End Get
        Set(ByVal value As DataSet)
            _ds = value
        End Set
    End Property
    Public Property ReportType() As String
        Get
            Return _sReportType
        End Get
        Set(ByVal value As String)
            _sReportType = value
        End Set
    End Property
    Public Property StartingItemCode() As String
        Get
            Return _sItemFr
        End Get
        Set(ByVal value As String)
            _sItemFr = value
        End Set
    End Property
    Public Property EndingItemCode() As String
        Get
            Return _sItemTo
        End Get
        Set(ByVal value As String)
            _sItemTo = value
        End Set
    End Property
    Public Property StartingItemType() As String
        Get
            Return _sItemTypeFr
        End Get
        Set(ByVal value As String)
            _sItemTypeFr = value
        End Set
    End Property
    Public Property EndingItemType() As String
        Get
            Return _sItemTypeTo
        End Get
        Set(ByVal value As String)
            _sItemTypeTo = value
        End Set
    End Property
    Public Property StartingItemGroup() As String
        Get
            Return _sItemGrpFr
        End Get
        Set(ByVal value As String)
            _sItemGrpFr = value
        End Set
    End Property
    Public Property EndingItemGroup() As String
        Get
            Return _sItemGrpTo
        End Get
        Set(ByVal value As String)
            _sItemGrpTo = value
        End Set
    End Property
    Public Property StartingWarehouse() As String
        Get
            Return _sWhseFr
        End Get
        Set(ByVal value As String)
            _sWhseFr = value
        End Set
    End Property
    Public Property EndingWarehouse() As String
        Get
            Return _sWhseTo
        End Get
        Set(ByVal value As String)
            _sWhseTo = value
        End Set
    End Property
    Public Property StartingDim1() As String
        Get
            Return _sDim1Fr
        End Get
        Set(ByVal value As String)
            _sDim1Fr = value
        End Set
    End Property
    Public Property EndingDim1() As String
        Get
            Return _sDim1To
        End Get
        Set(ByVal value As String)
            _sDim1To = value
        End Set
    End Property
    Public Property StartingDim2() As String
        Get
            Return _sDim2Fr
        End Get
        Set(ByVal value As String)
            _sDim2Fr = value
        End Set
    End Property
    Public Property EndingDim2() As String
        Get
            Return _sDim2To
        End Get
        Set(ByVal value As String)
            _sDim2To = value
        End Set
    End Property
    Public Property AsAtDate() As DateTime
        Get
            Return _AsAtDate
        End Get
        Set(ByVal value As DateTime)
            _AsAtDate = value
        End Set
    End Property
    Public Property LocalCurrency() As String
        Get
            Return _LocalCurr
        End Get
        Set(ByVal value As String)
            _LocalCurr = value
        End Set
    End Property
    Public Property IsReportShared() As Boolean
        Get
            Return _isRptShared
        End Get
        Set(ByVal value As Boolean)
            _isRptShared = value
        End Set
    End Property
    Public Property CrystalReportPath() As String
        Get
            Return _sRptPath
        End Get
        Set(ByVal value As String)
            _sRptPath = value
        End Set
    End Property
#End Region

    Private Sub OpenStockAging()
        Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
        Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
        Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo
        Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
        Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
        Dim rpt As CrystalDecisions.CrystalReports.Engine.ReportDocument

        crConnectionInfo.ServerName = oCompany.Server
        crConnectionInfo.DatabaseName = oCompany.CompanyDB
        crConnectionInfo.UserID = DBUsername
        crConnectionInfo.Password = DBPassword

        Select Case _sReportType
            Case "0"
                If _isRptShared Then
                    rpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                    rpt.Load(_sRptPath)
                Else
                    rpt = New rptTM_SAR
                End If
            Case "1"
                If _isRptShared Then
                    rpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                    rpt.Load(_sRptPath)
                Else
                    rpt = New rptTM_SAR_V2
                End If
            Case "2"
                If _isRptShared Then
                    rpt = New CrystalDecisions.CrystalReports.Engine.ReportDocument
                    rpt.Load(_sRptPath)
                Else
                    rpt = New rptTM_SAR_V3
                End If
        End Select
        
        With rpt
            Dim dtTemp As DateTime = DateTime.Now

            rpt.DataDefinition.FormulaFields.Item("Username").Text = "'" & oCompany.UserName & "'"
            rpt.DataDefinition.FormulaFields.Item("Company").Text = "'" & oCompany.CompanyName & "'"
            rpt.DataDefinition.FormulaFields.Item("LocalCurrency").Text = "'" & _LocalCurr & "'"
            rpt.DataDefinition.FormulaFields.Item("AsAtDate").Text = "Date(" & AsAtDate.Year & ", " & AsAtDate.Month.ToString("00") & ", " & AsAtDate.Day.ToString("00") & ") "
            rpt.DataDefinition.FormulaFields.Item("StartingItemCode").Text = "'" & _sItemFr & "'"
            rpt.DataDefinition.FormulaFields.Item("EndingItemCode").Text = "'" & _sItemTo & "'"
            rpt.DataDefinition.FormulaFields.Item("StartingItemGrp").Text = "'" & _sItemGrpFr & "'"
            rpt.DataDefinition.FormulaFields.Item("EndingItemGrp").Text = "'" & _sItemGrpTo & "'"
            rpt.DataDefinition.FormulaFields.Item("StartingWhseCode").Text = "'" & _sWhseFr & "'"
            rpt.DataDefinition.FormulaFields.Item("EndingWhseCode").Text = "'" & _sWhseTo & "'"
            rpt.DataDefinition.FormulaFields.Item("StartingDim1").Text = "'" & _sDim1Fr & "'"
            rpt.DataDefinition.FormulaFields.Item("EndingDim1").Text = "'" & _sDim1To & "'"
            rpt.DataDefinition.FormulaFields.Item("StartingDim2").Text = "'" & _sDim2Fr & "'"
            rpt.DataDefinition.FormulaFields.Item("EndingDim2").Text = "'" & _sDim2To & "'"
            rpt.DataDefinition.FormulaFields.Item("StartingItemType").Text = "'" & _sItemTypeFr & "'"
            rpt.DataDefinition.FormulaFields.Item("EndingItemType").Text = "'" & _sItemTypeTo & "'"

            Dim bLoadDatasource As Boolean = False

            CrTables = .Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
                CrTable.Location = oCompany.CompanyDB & ".dbo." & CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
                bLoadDatasource = True

            Next


            .RecordSelectionFormula = "{_NCM_TM_SARY.Username} = '" & oCompany.UserName & "'"
            .Refresh()
        End With
        crViewer.ReportSource = rpt
        crViewer.Show()
    End Sub

    Private Sub crViewer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles crViewer.Load
        Try
            OpenStockAging()
        Catch ex As Exception
            MessageBox.Show("[frmTMSAR_2].[crViewer_Load] - " & ex.Message, String.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        End Try
    End Sub
End Class