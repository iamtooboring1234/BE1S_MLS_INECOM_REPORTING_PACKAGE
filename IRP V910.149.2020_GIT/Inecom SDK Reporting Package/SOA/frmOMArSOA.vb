Imports System.IO
Imports System.Data.SqlClient
Imports SAPbobsCOM
Public Class frmOMArSOA

#Region "Global Variables"
    Private BPCode As String
    Private AsAtDate As DateTime
    Private FromDate As DateTime
    Private IsBBF As String
    Private IsGAT As String
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False

    ' IMPORTANT! Choose the correct company before compiling
    Private Const ClientCompany As CompanyCode = CompanyCode.General
    Private Const EmbeddedType As Boolean = False

    Private oFormARSOA As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oCheck As SAPbouiCOM.CheckBox
#End Region

#Region "Intialize Application"
    Public Sub New()
        Try
            'If Not NotesSetup() Then
            '    MsgBox("Error creating Database '@NCM_SOC2'")
            '    Exit Sub
            'End If
        Catch ex As Exception
            MsgBox("[frmOMArSoa].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Private Function setChooseFromListConditions(ByVal myVal As String, ByVal compareVal As String, ByVal cflUID As String) As Boolean
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        oCFL = oFormARSOA.ChooseFromLists.Item(cflUID)
        oCons = New SAPbouiCOM.Conditions
        Try
            Select Case cflUID
                Case "cflBPFr"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    End If

                    Exit Select
                Case "cflBPTo"
                    If (myVal.Length > 0 AndAlso compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (myVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCon.CondVal = myVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    ElseIf (compareVal.Length > 0) Then
                        oCon = oCons.Add()
                        oCon.Alias = "CardCode"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL
                        oCon.CondVal = compareVal
                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                    End If
                    Exit Select
                Case Else
                    Throw New Exception("Invalid Choose from list. UID#" & cflUID)
                    Exit Select
            End Select
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            Return True
        Catch ex As Exception
            Throw New Exception("[frmOMArSoa].[SetChooseFromList] " & vbNewLine & ex.Message)
            Return False
        End Try
        Return False
    End Function
    Public Sub LoadForm()
        Dim oItem As SAPbouiCOM.Item
        Dim oPictureBox As SAPbouiCOM.PictureBox

        If LoadFromXML("Inecom_SDK_Reporting_Package.NCM_OMARSOA.srf") Then
            oFormARSOA = SBO_Application.Forms.Item("NCM_OMARSOA")
            oPictureBox = oFormARSOA.Items.Item("pbInecom").Specific
            oPictureBox.Picture = Application.StartupPath.ToString & "\ncmInecom.bmp"
            For Each oItem In oFormARSOA.Items
                oItem.FontSize = 10
            Next
            If ClientCompany = CompanyCode.AE Then
                oFormARSOA.Items.Item("ckLogo").Visible = False
            End If
            oFormARSOA.Items.Item("lbStatus").FontSize = 10
            oFormARSOA.Items.Item("lbStyleOpt").TextStyle = 4
            AddDataSource()
            SetDatasource()
            RetrieveNotes()
            SetupChooseFromList()
            oFormARSOA.Visible = True
        Else
            Try
                If oFormARSOA.Visible = False Then
                    oFormARSOA.Close()
                Else
                    oFormARSOA.Select()
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub
    Private Sub AddDataSource()
        Try
            With oFormARSOA.DataSources.UserDataSources
                .Add("DateType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("DateAsAt", SAPbouiCOM.BoDataType.dt_DATE)
                .Add("BPCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                .Add("Period", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("Logo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("HDR", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("BBF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("SNP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("GAT", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("HAS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("HFN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("EXC", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("Notes", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1000)
                .Add("txtBPFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtBPTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20)
                .Add("txtBPGFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtBPGTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtSlsFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
                .Add("txtSlsTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100)
            End With
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[AddDataSource] : " & ex.Message)
        End Try
    End Sub
    Private Sub SetDatasource()
        Try
            oEdit = oFormARSOA.Items.Item("etBPCode").Specific
            oEdit.DataBind.SetBound(True, "", "BPCode")
            oEdit = oFormARSOA.Items.Item("etDateAsAt").Specific
            oEdit.DataBind.SetBound(True, "", "DateAsAt")
            oEdit.Value = Now.ToString("yyyyMMdd")
            oEdit = oFormARSOA.Items.Item("etNotes").Specific
            oEdit.DataBind.SetBound(True, "", "Notes")

            oEdit = oFormARSOA.Items.Item("txtBPFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPFr")
            oEdit = oFormARSOA.Items.Item("txtBPTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPTo")
            oEdit = oFormARSOA.Items.Item("txtBPGFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPGFr")
            oEdit = oFormARSOA.Items.Item("txtBPGTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtBPGTo")
            oEdit = oFormARSOA.Items.Item("txtSlsFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtSlsFr")
            oEdit = oFormARSOA.Items.Item("txtSlsTo").Specific
            oEdit.DataBind.SetBound(True, "", "txtSlsTo")


            oCombo = oFormARSOA.Items.Item("cbDateType").Specific
            oCombo.ValidValues.Add("0", "Document Date")
            oCombo.ValidValues.Add("1", "Due Date")
            oCombo.ValidValues.Add("2", "Posting Date")
            oCombo.DataBind.SetBound(True, "", "DateType")
            oFormARSOA.DataSources.UserDataSources.Item("DateType").ValueEx = "0"

            oCombo = oFormARSOA.Items.Item("cbPrdType").Specific
            oCombo.ValidValues.Add("0", "Every 30 Days")
            oCombo.ValidValues.Add("1", "Every Month")
            oCombo.DataBind.SetBound(True, "", "Period")
            oFormARSOA.DataSources.UserDataSources.Item("Period").ValueEx = "0"

            oCheck = oFormARSOA.Items.Item("ckLogo").Specific
            oCheck.DataBind.SetBound(True, "", "Logo")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            If ClientCompany = CompanyCode.AMS Then
                oFormARSOA.Items.Item("ckLogo").Enabled = False
            End If

            oCheck = oFormARSOA.Items.Item("ckHDR").Specific
            oCheck.DataBind.SetBound(True, "", "HDR")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
            oCheck.DataBind.SetBound(True, "", "BBF")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckSNP").Specific
            oCheck.DataBind.SetBound(True, "", "SNP")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
            oCheck.DataBind.SetBound(True, "", "GAT")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckHAS").Specific
            oCheck.DataBind.SetBound(True, "", "HAS")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckHFN").Specific
            oCheck.DataBind.SetBound(True, "", "HFN")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
            oCheck = oFormARSOA.Items.Item("ckExc").Specific
            oCheck.DataBind.SetBound(True, "", "EXC")
            oCheck.ValOff = "N"
            oCheck.ValOn = "Y"
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[SetDatasource] : " & ex.Message)
        End Try
    End Sub
    Private Sub SetupChooseFromList()
        Dim oEditLn As SAPbouiCOM.EditText
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFLCreation As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oFormARSOA.ChooseFromLists
            'Production Order Choose from list
            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "2"
            oCFLCreation.UniqueID = "cflBPFr"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtBPFr").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflBPFr"
            oEditLn.ChooseFromListAlias = "CardCode"

            oCFLCreation = DirectCast(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams), SAPbouiCOM.ChooseFromListCreationParams)
            oCFLCreation.MultiSelection = False
            oCFLCreation.ObjectType = "2"
            oCFLCreation.UniqueID = "cflBPTo"
            oCFL = oCFLs.Add(oCFLCreation)

            oEditLn = DirectCast(oFormARSOA.Items.Item("txtBPTo").Specific, SAPbouiCOM.EditText)
            oEditLn.ChooseFromListUID = "cflBPTo"
            oEditLn.ChooseFromListAlias = "CardCode"

        Catch ex As Exception
            Throw New Exception("[frmOMArSoa].[SetupChooseFromList]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Sub ShowStatus(ByVal sStatus As String)
        Try
            Dim oStaticText As SAPbouiCOM.StaticText = oFormARSOA.Items.Item("lbStatus").Specific
            oStaticText.Caption = sStatus
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[ShowStatus] : " & ex.Message)
        End Try
    End Sub
    Private Sub LoadViewer()
        Try
            Dim frm As New Hydac_FormViewer

            oCombo = oFormARSOA.Items.Item("cbDateType").Specific
            If oCombo.Selected Is Nothing Then
                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            frm.Report = oCombo.Selected.Value

            oCombo = oFormARSOA.Items.Item("cbPrdType").Specific
            If oCombo.Selected Is Nothing Then
                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

            frm.IsShared = g_bIsShared
            frm.SharedReportName = g_sReportFilename
            frm.Period = oCombo.Selected.Value
            frm.DBUsernameViewer = DBUsername
            frm.DBPasswordViewer = DBPassword
            frm.Username = oCompany.UserName
            frm.AsAtDate = AsAtDate.ToString("yyyyMMdd")

            oCheck = oFormARSOA.Items.Item("ckLogo").Specific
            frm.HideLogo = IIf(oCheck.Checked, True, False)
            oCheck = oFormARSOA.Items.Item("ckHDR").Specific
            frm.HideHeader = IIf(oCheck.Checked, True, False)
            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
            frm.IsBBF = IIf(oCheck.Checked, 1, 0)
            oCheck = oFormARSOA.Items.Item("ckSNP").Specific
            frm.IsSNP = IIf(oCheck.Checked, 1, 0)
            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
            frm.IsGAT = IIf(oCheck.Checked, 1, 0)
            oCheck = oFormARSOA.Items.Item("ckHAS").Specific
            frm.IsHAS = IIf(oCheck.Checked, 1, 0)
            oCheck = oFormARSOA.Items.Item("ckHFN").Specific
            frm.IsHFN = IIf(oCheck.Checked, 1, 0)
            frm.ReportName = ReportName.OMARSoa
            frm.LocalCurrency = GetLocalCurrency()
            frm.CompanySOA = ClientCompany
            frm.ShowDialog()

        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[LoadViewer] : " & ex.Message)
        End Try
    End Sub
    Private Function IsSharedFileExist() As Boolean
        Try
            g_sReportFilename = GetSharedFilePath(ReportName.OMARSoa)
            If g_sReportFilename <> "" Then
                If IsSharedFilePathExists(g_sReportFilename) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[AR SOA - OM].[GetPath] :" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
#End Region

#Region "Logic Function"
    Private Function NotesSetup() As Boolean
        Dim bSuccess As Boolean = False
        Dim SQLCommand As System.Data.SqlClient.SqlCommand
        Dim sQuery1 As String = "IF NOT EXISTS(SELECT name FROM " & oCompany.CompanyDB & ".dbo.sysobjects WHERE xtype = 'U' AND name = '@NCM_SOC2')"
        sQuery1 &= vbNewLine & " BEGIN "
        sQuery1 &= vbNewLine & " USE " & oCompany.CompanyDB
        sQuery1 &= vbNewLine & " CREATE TABLE [@NCM_SOC2]"
        sQuery1 &= vbNewLine & " (ID         NVARCHAR(8)         NOT NULL,"
        sQuery1 &= vbNewLine & " Notes      NVARCHAR(1000)      NOT NULL,"
        sQuery1 &= vbNewLine & " Image    IMAGE)"
        sQuery1 &= vbNewLine & " INSERT INTO [@NCM_SOC2]"
        sQuery1 &= vbNewLine & " VALUES ("
        sQuery1 &= vbNewLine & " '1',"
        sQuery1 &= vbNewLine & " 'Note:   Any payments received after end of the month will be shown in next month''s statement."
        sQuery1 &= vbNewLine & "           If you do not agree with the above statement, please inform us immediately.'"
        sQuery1 &= vbNewLine & " , NULL)"
        sQuery1 &= vbNewLine & " END"

        Try
            SQLCommand = SQLDbConnection.CreateCommand
            SQLCommand.CommandText = sQuery1
            SQLCommand.CommandType = CommandType.Text
            SQLCommand.ExecuteNonQuery()
            bSuccess = True
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[NoteSetups]" & vbNewLine & ex.Message)
            bSuccess = False
        End Try
        Return bSuccess
    End Function
    Private Sub RetrieveNotes()
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery("SELECT Notes FROM [@NCM_SOC2] WHERE ID='1'")
            If oRecord.RecordCount > 0 Then
                oFormARSOA.DataSources.UserDataSources.Item("Notes").ValueEx = oRecord.Fields.Item(0).Value
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[RetrieveNotes]" & vbNewLine & ex.Message)
        End Try
    End Sub
    Private Function SaveSettings() As Boolean
        Dim Notes As String
        Dim BitmapPath As String
        Dim ImagePath As String = ""
        Dim Image As Byte()
        Dim cmd As SqlCommand
        Dim sQuery As String
        Dim FileStrm As FileStream
        Dim BinReader As BinaryReader

        Try
            ShowStatus("Status: Saving Settings...")
            Notes = oFormARSOA.DataSources.UserDataSources.Item("Notes").ValueEx
            Notes = Notes.Replace("'", "''")
            sQuery = "Select BitmapPath From OADP"
            cmd = New SqlCommand(sQuery, SQLDbConnection)
            BitmapPath = cmd.ExecuteScalar

            If ClientCompany <> CompanyCode.AMS Then
                ImagePath = BitmapPath & oCompany.CompanyDB & ".bmp"
                If File.Exists(ImagePath) = False Then
                    ImagePath = BitmapPath & oCompany.CompanyDB & ".jpg"
                    If File.Exists(ImagePath) = False Then
                        ImagePath = BitmapPath & oCompany.CompanyDB & ".png"
                        If File.Exists(ImagePath) = False Then
                            ImagePath = BitmapPath & oCompany.CompanyDB & ".tiff"
                            If File.Exists(ImagePath) = False Then
                                ImagePath = ""
                            End If
                        End If
                    End If
                End If
            End If

            'Read the file 
            If ImagePath.Trim <> "" Then
                FileStrm = New FileStream(ImagePath, FileMode.Open)
                BinReader = New BinaryReader(FileStrm)
                Image = BinReader.ReadBytes(BinReader.BaseStream.Length)
                FileStrm.Close()
                BinReader.Close()

                sQuery = "UPDATE [@NCM_SOC2] SET Notes='" & Notes & "', Image=@Image WHERE ID = '1'"
                cmd = New SqlCommand(sQuery, SQLDbConnection)
                cmd.Parameters.Add("@Image", Image)
                cmd.ExecuteNonQuery()
            Else
                sQuery = "UPDATE [@NCM_SOC2] SET Notes='" & Notes & "', Image=0x0 WHERE ID = '1'"
                cmd = New SqlCommand(sQuery, SQLDbConnection)
                cmd.ExecuteNonQuery()
            End If

            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[SaveImages]:" & ex.Message)
            Return False
        End Try
    End Function
    Private Function MainCtrlNONEMBEDDED() As Boolean

        g_bIsShared = IsSharedFileExist()
        If (g_bIsShared) Then
            If (Not File.Exists(g_sReportFilename)) Then
                SBO_Application.StatusBar.SetText("[frmOMArSOA]: Crystal Report file is not found in location (" & g_sReportFilename & ")", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                System.Threading.Thread.CurrentThread.Sleep(1300)
                Return False
            End If
        Else
            If (g_sReportFilename.Length = 1 AndAlso g_sReportFilename.Trim().Length = 0) Then
                Return False
            End If
        End If
        Dim sDate As String = String.Empty
        Dim sBBF As String = "N"
        Dim bSuccess As Boolean = False
        Dim iRowsAffected As Integer = 0
        Dim sQuery As String = String.Empty
        Dim SQLCommand As System.Data.SqlClient.SqlCommand

        Dim sBPCodeFr As String = String.Empty
        Dim sBPCodeTo As String = String.Empty
        Dim sBPGrpFr As String = String.Empty
        Dim sBPGrpTo As String = String.Empty
        Dim sSlsFr As String = String.Empty
        Dim sSlsTo As String = String.Empty
        Try
            'Get Parameter Value
            oEdit = oFormARSOA.Items.Item("txtBPFr").Specific
            sBPCodeFr = oEdit.Value

            oEdit = oFormARSOA.Items.Item("txtBPTo").Specific
            sBPCodeTo = oEdit.Value

            oEdit = oFormARSOA.Items.Item("txtBPGFr").Specific
            sBPGrpFr = oEdit.Value

            oEdit = oFormARSOA.Items.Item("txtBPGTo").Specific
            sBPGrpTo = oEdit.Value

            oEdit = oFormARSOA.Items.Item("txtSlsFr").Specific
            sSlsFr = oEdit.Value

            oEdit = oFormARSOA.Items.Item("txtSlsTo").Specific
            sSlsTo = oEdit.Value


            'Get BPCode
            oEdit = oFormARSOA.Items.Item("etBPCode").Specific
            BPCode = CType(IIf(oEdit.Value = "", "%", "%" & oEdit.Value.Replace("*", "%")) & "%", String).Trim

            'Get AsAtDate, FromDate
            oEdit = oFormARSOA.Items.Item("etDateAsAt").Specific
            sDate = oEdit.Value.Trim
            If sDate = "" Then Throw New Exception("Error: As At Date is empty!")
            AsAtDate = New DateTime(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
            FromDate = New DateTime(Left(sDate, 4), Mid(sDate, 5, 2), "01")

            'Get IsBBF
            oCheck = oFormARSOA.Items.Item("ckBBF").Specific
            If oCheck.Checked Then IsBBF = "Y" Else IsBBF = "N"

            'Get IsGAT
            oCheck = oFormARSOA.Items.Item("ckGAT").Specific
            If oCheck.Checked Then IsGAT = "Y" Else IsGAT = "N"

            'Set the query
            sQuery = "EXECUTE SP_SOA '"
            sQuery &= oCompany.UserName & "','"
            sQuery &= sBPCodeFr & "','"
            sQuery &= sBPCodeTo & "','"
            sQuery &= sBPGrpFr & "','"
            sQuery &= sBPGrpTo & "','"
            sQuery &= sSlsFr & "','"
            sQuery &= sSlsTo & "','"
            sQuery &= BPCode & "','"
            sQuery &= FromDate.ToString("yyyyMMdd") & "','"
            sQuery &= AsAtDate.ToString("yyyyMMdd") & "','"
            sQuery &= IsBBF & "','"
            sQuery &= IsGAT & "',"

            oCheck = oFormARSOA.Items.Item("ckExc").Specific
            If oCheck.Checked Then sQuery &= "1" Else sQuery &= "0"

            Try
                ShowStatus("Status: Executing Procedure...")
                Dim mySqlConn As New System.Data.SqlClient.SqlConnection(DBConnString)
                Try
                    mySqlConn.Open()
                Catch ex As Exception
                    Throw ex
                End Try
                Try
                    SQLCommand = SQLDbConnection.CreateCommand
                    SQLCommand.CommandTimeout = 3600
                    SQLCommand.CommandText = sQuery
                    SQLCommand.CommandType = CommandType.Text
                    SQLCommand.ExecuteNonQuery()
                Catch ex As Exception
                    Throw ex
                Finally
                    If mySqlConn.State = ConnectionState.Open Then
                        mySqlConn.Close()
                    End If
                End Try
                ShowStatus("Status: Completed!")
                bSuccess = True
            Catch ex As Exception
                bSuccess = False
                Throw ex
            End Try
            SBO_Application.StatusBar.SetText("Completed Successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[MainCtrlNONEMBEDDED]" & vbNewLine & ex.Message)
        End Try
        Return bSuccess
    End Function
    Private Function ValidateParameter() As Boolean

        Try
            oFormARSOA.ActiveItem = "etBPCode"
            Dim oRecordsetLn As SAPbobsCOM.Recordset
            oRecordsetLn = DirectCast(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            Dim sStart As String = String.Empty
            Dim sEnd As String = String.Empty
            Dim sQuery As String = String.Empty

            sStart = oFormARSOA.DataSources.UserDataSources.Item("txtBPFr").ValueEx
            sEnd = oFormARSOA.DataSources.UserDataSources.Item("txtBPTo").ValueEx
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("BP Code from is greater than BP Code to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtBPFr"
                    Return False
                End If
            End If


            sStart = oFormARSOA.DataSources.UserDataSources.Item("txtBPGFr").ValueEx
            sEnd = oFormARSOA.DataSources.UserDataSources.Item("txtBPGTo").ValueEx
            If (sStart.Length > 0) Then
                sQuery = "SELECT * FROM OCRG WHERE GROUPTYPE = 'C' AND GROUPCODE = '" & sStart & "'"
                oRecordsetLn.DoQuery(sQuery)
                If (oRecordsetLn.RecordCount = 0) Then
                    SBO_Application.StatusBar.SetText("Invalid BP Group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtBPGFr"
                    Return False
                End If
            End If

            If (sEnd.Length > 0) Then
                sQuery = "SELECT * FROM OCRG WHERE GROUPTYPE = 'C' AND GROUPCODE = '" & sEnd & "'"
                oRecordsetLn.DoQuery(sQuery)
                If (oRecordsetLn.RecordCount = 0) Then
                    SBO_Application.StatusBar.SetText("Invalid BP Group", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtBPGTo"
                    Return False
                End If
            End If

            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("BP Group from is greater than BP Group to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtBPGFr"
                    Return False
                End If
            End If

            sStart = oFormARSOA.DataSources.UserDataSources.Item("txtSlsFr").ValueEx
            sEnd = oFormARSOA.DataSources.UserDataSources.Item("txtSlsTo").ValueEx
            If (sStart.Length > 0) Then
                sQuery = "SELECT * FROM OSLP WHERE SLPNAME = '" & sStart & "'"
                oRecordsetLn.DoQuery(sQuery)
                If (oRecordsetLn.RecordCount = 0) Then
                    SBO_Application.StatusBar.SetText("Invalid Sales Employee", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtSlsFr"
                    Return False
                End If
            End If
            If (sEnd.Length > 0) Then
                sQuery = "SELECT * FROM OSLP WHERE SLPNAME = '" & sEnd & "'"
                oRecordsetLn.DoQuery(sQuery)
                If (oRecordsetLn.RecordCount = 0) Then
                    SBO_Application.StatusBar.SetText("Invalid Sales Employee", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtSlsTo"
                    Return False
                End If
            End If
            If (sStart.Length > 0 AndAlso sEnd.Length > 0) Then
                If (String.Compare(sStart, sEnd) > 0) Then
                    SBO_Application.StatusBar.SetText("Sales Employee from is greater than Sales Employee to", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oFormARSOA.ActiveItem = "txtSlsFr"
                    Return False
                End If
            End If
            SBO_Application.StatusBar.SetText(String.Empty, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None)
            Return True
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[ValidateParameter] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function GetLocalCurrency() As String
        Try
            Dim oRecordsetLn As SAPbobsCOM.Recordset
            oRecordsetLn = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQueryLn As String = String.Empty
            sQueryLn = "SELECT	DISTINCT MainCurncy FROM OADM "
            oRecordsetLn.DoQuery(sQueryLn)
            If (oRecordsetLn.RecordCount > 0) Then
                Return oRecordsetLn.Fields.Item(0).Value.ToString()
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[frmOMArSOA].[GetLocalCurrency] - " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return String.Empty
    End Function
#End Region

#Region "Events Handler"
    Public Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action = True Then
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then
                    Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvent = pVal
                    Dim myVal As String = String.Empty
                    Dim compareVal As String = String.Empty
                    Select Case oCFLEvent.ChooseFromListUID
                        Case "cflBPFr"
                            myVal = DirectCast(oFormARSOA.Items.Item("txtBPFr").Specific, SAPbouiCOM.EditText).Value
                            compareVal = String.Empty 'DirectCast(oFormARSOA.Items.Item("txtBPTo").Specific, SAPbouiCOM.EditText).Value
                            Exit Select
                        Case "cflBPTo"
                            myVal = DirectCast(oFormARSOA.Items.Item("txtBPTo").Specific, SAPbouiCOM.EditText).Value
                            compareVal = String.Empty 'DirectCast(oFormARSOA.Items.Item("txtBPFr").Specific, SAPbouiCOM.EditText).Value
                            Exit Select
                    End Select
                    Return setChooseFromListConditions(myVal, compareVal, oCFLEvent.ChooseFromListUID)
                End If
                Select Case pVal.ItemUID
                    Case "etBPCode", "txtBPGFr", "txtBPGTo", "txtSlsFr", "txtSlsTo"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                            oEdit = oFormARSOA.Items.Item(pVal.ItemUID).Specific
                            If (oEdit.Value = String.Empty) And (pVal.CharPressed = 9) Then
                                SBO_Application.SendKeys("+{F2}")
                                BubbleEvent = False
                            End If
                        End If
                        Exit Select
                    Case "ckHFN"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            oFormARSOA.Items.Item("etBPCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    Case "btnExecute"
                        Return ValidateParameter()
                End Select
            Else
                If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST) Then
                    Dim oCFLEvent As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvent = pVal
                    Dim oDataTable As SAPbouiCOM.DataTable
                    oDataTable = oCFLEvent.SelectedObjects
                    If (Not oDataTable Is Nothing) Then
                        Dim sTemp As String = String.Empty
                        Select Case oCFLEvent.ChooseFromListUID
                            Case "cflBPFr"
                                sTemp = oDataTable.GetValue("CardCode", 0)
                                oFormARSOA.DataSources.UserDataSources.Item("txtBPFr").ValueEx = sTemp
                                Exit Select
                            Case "cflBPTo"
                                sTemp = oDataTable.GetValue("CardCode", 0)
                                oFormARSOA.DataSources.UserDataSources.Item("txtBPTo").ValueEx = sTemp
                                Exit Select
                            Case Else
                                Exit Select
                        End Select
                        Return True
                    End If
                End If
                Select Case pVal.ItemUID
                    Case "btnExecute"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If SaveSettings() Then
                                If EmbeddedType Then

                                Else
                                    If MainCtrlNONEMBEDDED() Then
                                        Dim myThread As New System.Threading.Thread(AddressOf LoadViewer)
#If B1Version = 2007 Then
                                        myThread.SetApartmentState(Threading.ApartmentState.STA)
#End If
                                        myThread.Start()
                                    End If
                                End If
                            End If
                        End If
                    Case "ckHFN"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            oFormARSOA.Items.Item("etNotes").Enabled = Not (oFormARSOA.Items.Item("etNotes").Enabled)
                        End If
                End Select
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[frmOMArSoa].[ItemEvent]" & vbNewLine & ex.Message)
            BubbleEvent = False
        End Try
        Return BubbleEvent
    End Function
#End Region

End Class

