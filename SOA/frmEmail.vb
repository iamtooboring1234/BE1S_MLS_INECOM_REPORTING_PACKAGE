'' © Copyright © 2007-2020, Inecom Pte Ltd, All rights reserved.
'' =============================================================

Imports System.IO
Imports System.Net
Imports outlook = Microsoft.Office.Interop.Outlook

Public Class frmEmail

#Region "Global Variables"
    Private sQuery As String = ""
    Private BPCode As String
    Private AsAtDate As DateTime
    Private FromDate As DateTime
    Private IsBBF As String
    Private IsGAT As String
    Private g_sReportFilename As String = String.Empty
    Private g_bIsShared As Boolean = False

    Private oForm As SAPbouiCOM.Form
    Private oEdit As SAPbouiCOM.EditText
    Private oCombo As SAPbouiCOM.ComboBox
    Private oChck As SAPbouiCOM.CheckBox
    Private txtPass As SAPbouiCOM.EditText
    Private oMtrx As SAPbouiCOM.Matrix
    Private oFldr As SAPbouiCOM.Folder

    Private g_iPaneLvl As Integer = 1

#End Region

#Region "Intialize Application"
    Public Sub New()
        Try

        Catch ex As Exception
            MsgBox("[frmEmail].[New()]" & vbNewLine & ex.Message)
        End Try
    End Sub
#End Region

#Region "General Functions"
    Public Sub LoadForm()
        If LoadFromXML("Inecom_SDK_Reporting_Package.ncmSOA_Email.srf") Then
            oForm = SBO_Application.Forms.Item("ncmSOA_Email")
            oForm.SupportedModes = -1

            g_iPaneLvl = 1

            oForm.Items.Item("tbReceipt").AffectsFormMode = False
            oForm.Items.Item("flHTML").AffectsFormMode = False
            oForm.Items.Item("flSMTP").AffectsFormMode = False

            oMtrx = oForm.Items.Item("mxEmail").Specific
            oMtrx.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            AddDataSource()
            SetDatasource()

            If (IsIncludeModule(ReportName.AR_Invoice)) Then
                oForm.Items.Item("lbINV").Visible = True
                oForm.Items.Item("lbDPI").Visible = True
                oForm.Items.Item("lbRIN").Visible = True
                oForm.Items.Item("tbINV").Visible = True
                oForm.Items.Item("tbDPI").Visible = True
                oForm.Items.Item("tbRIN").Visible = True
            Else
                oForm.Items.Item("lbINV").Visible = False
                oForm.Items.Item("lbDPI").Visible = False
                oForm.Items.Item("lbRIN").Visible = False
                oForm.Items.Item("tbINV").Visible = False
                oForm.Items.Item("tbDPI").Visible = False
                oForm.Items.Item("tbRIN").Visible = False

                oForm.Items.Item("lbINV").FromPane = 3
                oForm.Items.Item("lbDPI").FromPane = 3
                oForm.Items.Item("lbRIN").FromPane = 3
                oForm.Items.Item("tbINV").FromPane = 3
                oForm.Items.Item("tbDPI").FromPane = 3
                oForm.Items.Item("tbRIN").FromPane = 3

                oForm.Items.Item("lbINV").ToPane = 3
                oForm.Items.Item("lbDPI").ToPane = 3
                oForm.Items.Item("lbRIN").ToPane = 3
                oForm.Items.Item("tbINV").ToPane = 3
                oForm.Items.Item("tbDPI").ToPane = 3
                oForm.Items.Item("tbRIN").ToPane = 3
            End If

            oForm.Title = "Email Configuration"
            oForm.PaneLevel = 1
            oForm.EnableMenu(MenuID.Add_Row, True)
            oForm.EnableMenu(MenuID.Delete_Row, True)

            oForm.Items.Item("flSMTP").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("txtMailFr").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Try
                oForm.Items.Item("btConnect").Visible = True
            Catch ex As Exception

            End Try

            oForm.Items.Item("tbRA").Visible = False
            oForm.Items.Item("tbPV").Visible = False
            oForm.Items.Item("lbRA").Visible = False
            oForm.Items.Item("lbPV").Visible = False

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            oForm.Visible = True

        Else
            Try
                If oForm.Visible = False Then
                    oForm.Close()
                Else
                    oForm.Select()
                End If
            Catch ex As Exception
                MessageBox.Show("[frmEmail].[LoadForm] - " & ex.Message)
            End Try
        End If
    End Sub
    Private Sub AddDataSource()
        Try

            Dim oColn As SAPbouiCOM.Column

            With oForm.DataSources.UserDataSources
                .Add("txtMailFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("txtSMTP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("cboAuth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 2)
                .Add("txtUser", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("txtPass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbMailBody", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbPortNum", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
                .Add("tbReceipt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("ckOffice", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("ckOutlook", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

                .Add("flSMTP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)
                .Add("flHTML", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1)

                .Add("tbARSOA", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbINV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbRIN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbDPI", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbPV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)
                .Add("tbRA", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254)

                .Add("cRow", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 10)
                .Add("cCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30)
                .Add("cEmail", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 1000)
            End With

            oMtrx = oForm.Items.Item("mxEmail").Specific

            oColn = oMtrx.Columns.Item("cRow")
            oColn.DataBind.SetBound(True, "", "cRow")
            oColn = oMtrx.Columns.Item("cEmail")
            oColn.DataBind.SetBound(True, "", "cEmail")
            oColn = oMtrx.Columns.Item("cCode")
            oColn.DataBind.SetBound(True, "", "cCode")
            oColn.Visible = False

            oFldr = oForm.Items.Item("flSMTP").Specific
            oFldr.DataBind.SetBound(True, "", "flSMTP")
            oFldr = oForm.Items.Item("flHTML").Specific
            oFldr.DataBind.SetBound(True, "", "flHTML")
            oFldr.GroupWith("flSMTP")

            oEdit = oForm.Items.Item("tbARSOA").Specific
            oEdit.DataBind.SetBound(True, "", "tbARSOA")
            oEdit = oForm.Items.Item("tbINV").Specific
            oEdit.DataBind.SetBound(True, "", "tbINV")
            oEdit = oForm.Items.Item("tbRIN").Specific
            oEdit.DataBind.SetBound(True, "", "tbRIN")
            oEdit = oForm.Items.Item("tbDPI").Specific
            oEdit.DataBind.SetBound(True, "", "tbDPI")
            oEdit = oForm.Items.Item("tbPV").Specific
            oEdit.DataBind.SetBound(True, "", "tbPV")
            oEdit = oForm.Items.Item("tbRA").Specific
            oEdit.DataBind.SetBound(True, "", "tbRA")

            oEdit = oForm.Items.Item("tbReceipt").Specific
            oEdit.DataBind.SetBound(True, "", "tbReceipt")
            'oEdit = oForm.Items.Item("tbMailBody").Specific
            'oEdit.DataBind.SetBound(True, "", "tbMailBody")
            oEdit = oForm.Items.Item("tbPortNum").Specific
            oEdit.DataBind.SetBound(True, "", "tbPortNum")
            oEdit = oForm.Items.Item("txtMailFr").Specific
            oEdit.DataBind.SetBound(True, "", "txtMailFr")
            oEdit = oForm.Items.Item("txtSMTP").Specific
            oEdit.DataBind.SetBound(True, "", "txtSMTP")
            oEdit = oForm.Items.Item("txtUser").Specific
            oEdit.DataBind.SetBound(True, "", "txtUser")
            oEdit = oForm.Items.Item("txtPass").Specific
            oEdit.DataBind.SetBound(True, "", "txtPass")

            oCombo = oForm.Items.Item("cboAuth").Specific
            oCombo.DataBind.SetBound(True, "", "cboAuth")

            oChck = oForm.Items.Item("ckOffice").Specific
            oChck.DataBind.SetBound(True, "", "ckOffice")
            oChck.ValOff = "N"
            oChck.ValOn = "Y"

            oChck = oForm.Items.Item("ckOutlook").Specific
            oChck.DataBind.SetBound(True, "", "ckOutlook")
            oChck.ValOff = "N"
            oChck.ValOn = "Y"
        Catch ex As Exception
            SBO_Application.StatusBar.SetSystemMessage("[Email - AddDataSource] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try
    End Sub
    Private Sub SetDatasource()
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty
            Dim sOutlook As String = "N"

            sQuery = "  SELECT IFNULL(T1.""FldValue"",''), IFNULL(T1.""Descr"",'') "
            sQuery &= " FROM ""CUFD"" T0 "
            sQuery &= " LEFT OUTER JOIN ""UFD1"" T1 "
            sQuery &= " ON T0.""TableID"" = T1.""TableID"" AND T0.""FieldID"" = T1.""FieldID"" "
            sQuery &= " WHERE T0.""TableID"" = '@NCM_EMAIL_CONFIG' AND T0.""AliasID"" = 'AuthType' "

            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                While Not oRecord.EoF
                    oCombo.ValidValues.Add(oRecord.Fields.Item(0).Value, oRecord.Fields.Item(1).Value)
                    oRecord.MoveNext()
                End While
            End If
            If oCombo.ValidValues.Count > 0 Then
                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End If

            Try
                sQuery = "  SELECT IFNULL(""U_Outlook"",'N') "
                sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCM_EMAIL_CONFIG"" "
                sQuery &= " WHERE ""Code"" = 'SOA' "
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)
                If oRecord.RecordCount > 0 Then
                    oRecord.MoveFirst()
                    sOutlook = oRecord.Fields.Item(0).Value.ToString.Trim.ToUpper
                End If
            Catch ex As Exception

            End Try

            sQuery = "  SELECT IFNULL(""U_AuthType"",'0'), "
            sQuery &= " IFNULL(""U_MailFrom"",''), "
            sQuery &= " IFNULL(""U_SMTP"",''), "
            sQuery &= " IFNULL(""U_Username"",''), "
            sQuery &= " IFNULL(""U_Password"",''), "
            sQuery &= " IFNULL(""U_EmailPath"",''), "
            sQuery &= " IFNULL(""U_PortNum"",0), "
            sQuery &= " IFNULL(""U_Office365"",'N') "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCM_EMAIL_CONFIG"" "
            sQuery &= " WHERE ""Code"" = 'SOA' "

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                With oForm.DataSources.UserDataSources
                    .Item("cboAuth").ValueEx = oRecord.Fields.Item(0).Value
                    .Item("txtMailFr").ValueEx = oRecord.Fields.Item(1).Value
                    .Item("txtSMTP").ValueEx = oRecord.Fields.Item(2).Value
                    .Item("txtUser").ValueEx = oRecord.Fields.Item(3).Value
                    .Item("txtPass").ValueEx = oRecord.Fields.Item(4).Value
                    .Item("tbARSOA").ValueEx = oRecord.Fields.Item(5).Value
                    .Item("tbPortNum").ValueEx = oRecord.Fields.Item(6).Value
                    .Item("ckOffice").ValueEx = oRecord.Fields.Item(7).Value
                    .Item("ckOutlook").ValueEx = sOutlook

                    If .Item("cboAuth").ValueEx.Trim() = "0" Then
                        oForm.Items.Item("txtUser").Enabled = False
                        oForm.Items.Item("txtPass").Enabled = False
                    Else
                        oForm.Items.Item("txtUser").Enabled = True
                        oForm.Items.Item("txtPass").Enabled = True
                    End If
                End With
            End If

            sQuery = "  SELECT TOP 1 IFNULL(""U_ARINV"",''), IFNULL(""U_ARRIN"",''), IFNULL(""U_ARDPI"",''), IFNULL(""U_PAYRA"",'') "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCM_EMAIL_HTML"" "

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                With oForm.DataSources.UserDataSources
                    .Item("tbINV").ValueEx = oRecord.Fields.Item(0).Value
                    .Item("tbRIN").ValueEx = oRecord.Fields.Item(1).Value
                    .Item("tbDPI").ValueEx = oRecord.Fields.Item(2).Value
                    .Item("tbRA").ValueEx = oRecord.Fields.Item(3).Value
                End With
            End If

            Dim iRow As Integer = 1

            sQuery = "  SELECT ""Code"", IFNULL(""U_EmailAdd"",'') "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCMCCEMAIL"" "
            sQuery &= " ORDER BY ""Code"" "

            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                oMtrx.Clear()

                While Not oRecord.EoF
                    With oForm.DataSources.UserDataSources
                        .Item("cRow").ValueEx = iRow
                        .Item("cCode").ValueEx = oRecord.Fields.Item(0).Value.ToString.Trim
                        .Item("cEmail").ValueEx = oRecord.Fields.Item(1).Value.ToString.Trim
                    End With
                    iRow += 1
                    oMtrx.AddRow()
                    oRecord.MoveNext()
                End While
                oMtrx.AutoResizeColumns()
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetSystemMessage("[Email - SetDataSources]  " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Logic Function"
    Private Function ValidateBeforeSave() As Boolean
        Try
            With oForm.DataSources.UserDataSources
                If .Item("txtMailFr").ValueEx.Trim.Length = 0 Then
                    oForm.ActiveItem = "txtMailFr"
                    Return False
                End If
                If .Item("txtSMTP").ValueEx.Trim.Length = 0 Then
                    oForm.ActiveItem = "txtSMTP"
                    Return False
                End If
                If .Item("cboAuth").ValueEx.Trim <> "0" Then
                    If .Item("txtUser").ValueEx.Trim.Length = 0 Then
                        oForm.ActiveItem = "txtUser"
                        Return False
                    End If
                    If .Item("txtPass").ValueEx.Trim.Length = 0 Then
                        oForm.ActiveItem = "txtPass"
                        Return False
                    End If
                End If

                If .Item("tbPortNum").ValueEx.ToString.Trim.Length > 0 Then
                    If Convert.ToInt32(.Item("tbPortNum").ValueEx.ToString) < 0 Then
                        SBO_Application.StatusBar.SetText("Port Num Is Not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbPortNum"
                        Return False
                    End If
                End If

                If .Item("tbARSOA").ValueEx.Trim.Length > 0 Then
                    If Not System.IO.File.Exists(.Item("tbARSOA").ValueEx.Trim) Then
                        SBO_Application.StatusBar.SetText("HTML File location for AR SOA Is Not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbARSOA"
                        Return False
                    End If
                End If

                If .Item("tbINV").ValueEx.Trim.Length > 0 Then
                    If Not System.IO.File.Exists(.Item("tbINV").ValueEx.Trim) Then
                        SBO_Application.StatusBar.SetText("HTML File location for AR Invoice Is Not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbINV"
                        Return False
                    End If
                End If

                If .Item("tbRIN").ValueEx.Trim.Length > 0 Then
                    If Not System.IO.File.Exists(.Item("tbRIN").ValueEx.Trim) Then
                        SBO_Application.StatusBar.SetText("HTML File location for AR Credit Note Is Not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbRIN"
                        Return False
                    End If
                End If

                If .Item("tbDPI").ValueEx.Trim.Length > 0 Then
                    If Not System.IO.File.Exists(.Item("tbDPI").ValueEx.Trim) Then
                        SBO_Application.StatusBar.SetText("HTML File location for AR DP Invoice Is Not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbDPI"
                        Return False
                    End If
                End If

                If .Item("tbRA").ValueEx.Trim.Length > 0 Then
                    If Not System.IO.File.Exists(.Item("tbRA").ValueEx.Trim) Then
                        SBO_Application.StatusBar.SetText("HTML File location for Remittance Advice Is Not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbRA"
                        Return False
                    End If
                End If

            End With
            Return True
        Catch ex As Exception
            Throw New Exception("[ValidateBeforeSave]" & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Private Function Save() As Boolean
        Dim oRecord As SAPbobsCOM.Recordset = Nothing
        Dim sQuery As String = ""
        Try
            Try
                sQuery = "DELETE FROM """ & oCompany.CompanyDB & """.""@NCM_EMAIL_CONFIG"" WHERE ""Code"" = 'SOA' "
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)

                sQuery = "INSERT INTO """ & oCompany.CompanyDB & """.""@NCM_EMAIL_CONFIG"" (""Code"", ""Name"", ""U_Type"", ""U_MailFrom"", ""U_SMTP"", ""U_Username"", ""U_Password"", ""U_AuthType"", ""U_PortNum"", ""U_EmailPath"",""U_Office365"") VALUES ('SOA', 'SOA', 'SOA', '{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}') "
                Dim sInput() As String = New String(7) {}
                With oForm.DataSources.UserDataSources
                    sInput(0) = .Item("txtMailFr").ValueEx
                    sInput(1) = .Item("txtSMTP").ValueEx
                    sInput(2) = .Item("txtUser").ValueEx
                    sInput(3) = .Item("txtPass").ValueEx
                    sInput(4) = .Item("cboAuth").ValueEx
                    sInput(5) = .Item("tbPortNum").ValueEx
                    sInput(6) = .Item("tbARSOA").ValueEx
                    sInput(7) = .Item("ckOffice").ValueEx
                End With

                sQuery = String.Format(sQuery, sInput)
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)

                Try
                    sQuery = "  UPDATE """ & oCompany.CompanyDB & """.""@NCM_EMAIL_CONFIG"" "
                    sQuery &= " SET ""U_Outlook"" = '" & oForm.DataSources.UserDataSources.Item("ckOutlook").ValueEx & "' "
                    sQuery &= " WHERE ""Code"" = 'SOA' "
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(sQuery)
                Catch ex As Exception

                End Try

                Try
                    sQuery = "  UPDATE """ & oCompany.CompanyDB & """.""@NCM_EMAIL_CONFIG"" "
                    sQuery &= " SET ""U_EnableSSL"" = '" & oForm.DataSources.UserDataSources.Item("ckOffice").ValueEx & "' "
                    sQuery &= " WHERE ""Code"" = 'SOA' "
                    oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecord.DoQuery(sQuery)
                Catch ex As Exception

                End Try

                sQuery = "DELETE FROM """ & oCompany.CompanyDB & """.""@NCM_EMAIL_HTML"" "
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)

                sQuery = "INSERT INTO """ & oCompany.CompanyDB & """.""@NCM_EMAIL_HTML"" (""U_ARSOA"", ""U_PAYPV"", ""U_ARINV"", ""U_ARRIN"", ""U_ARDPI"", ""U_PAYRA"") VALUES ('{0}','{1}','{2}','{3}','{4}','{5}') "
                Dim sInput2() As String = New String(6) {}
                With oForm.DataSources.UserDataSources
                    sInput2(0) = ""
                    sInput2(1) = ""
                    sInput2(2) = .Item("tbINV").ValueEx
                    sInput2(3) = .Item("tbRIN").ValueEx
                    sInput2(4) = .Item("tbDPI").ValueEx
                    sInput2(5) = .Item("tbRA").ValueEx
                End With

                sQuery = String.Format(sQuery, sInput2)
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)

                ' CC EMAIL
                ' -------------------------------------------------------------------------
                sQuery = "DELETE FROM """ & oCompany.CompanyDB & """.""@NCMCCEMAIL"" "
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)

                For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                    oMtrx.GetLineData(i)
                    If oForm.DataSources.UserDataSources.Item("cEmail").ValueEx.ToString.Trim <> "" Then
                        Try
                            sQuery = "  INSERT INTO """ & oCompany.CompanyDB & """.""@NCMCCEMAIL"" "
                            sQuery &= " (""Code"", ""Name"", ""U_EmailAdd"") VALUES "
                            sQuery &= " ('" & oForm.DataSources.UserDataSources.Item("cRow").ValueEx & "', "
                            sQuery &= "  '" & oForm.DataSources.UserDataSources.Item("cRow").ValueEx & "', "
                            sQuery &= "  '" & oForm.DataSources.UserDataSources.Item("cEmail").ValueEx & "') "

                            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecord.DoQuery(sQuery)
                        Catch ex As Exception

                        End Try
                    End If
                Next
                ' -------------------------------------------------------------------------

                oRecord = Nothing

                SBO_Application.StatusBar.SetText("The configuration has been updated successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            Catch ex As Exception
                Throw ex
            End Try
            Return True
        Catch ex As Exception
            Throw New Exception("[frmEmail].[Save]" & vbNewLine & ex.Message)
            Return False
        End Try
    End Function

    Private Function ValidateData() As Boolean
        Try
            For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                oMtrx.GetLineData(i)
                If oForm.DataSources.UserDataSources.Item("cEmail").ValueEx.ToString.Trim <> "" Then
                    If Not oForm.DataSources.UserDataSources.Item("cEmail").ValueEx.ToString.Trim.Contains("@") Then
                        SBO_Application.StatusBar.SetText("Line " & i & " - Invalid Email Address. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function CheckImage1(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image1") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage1] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage2(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image2") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage2] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage3(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image3") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage3] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage4(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image4") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage4] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage5(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image5") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage5] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage6(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image6") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage6] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function

    Private Sub Connect()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sRec As String = ""
            Dim sEnableSSL As String = ""
            Dim sToday As String = ""
            Dim sQry As String = "SELECT TO_VARCHAR(current_timestamp, 'DD.MM.YYYY HH:MM:SS') FROM DUMMY"

            sRec = " SELECT TOP 1 IFNULL(""U_EnableSSL"",'N') FROM ""@NCM_EMAIL_CONFIG"" WHERE ""Code"" = 'SOA' "
            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery(sRec)
            If oRec.RecordCount > 0 Then
                sEnableSSL = oRec.Fields.Item(0).Value
            End If

            Try
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQry)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    sToday = oRec.Fields.Item(0).Value.ToString.Trim
                End If
            Catch ex As Exception

            End Try

            Dim sInput() As String = New String(7) {}
            With oForm.DataSources.UserDataSources
                sInput(0) = .Item("txtMailFr").ValueEx.ToString.Trim
                sInput(1) = .Item("txtSMTP").ValueEx.ToString.Trim
                sInput(2) = .Item("txtUser").ValueEx.ToString.Trim
                sInput(3) = .Item("txtPass").ValueEx.ToString.Trim
                sInput(4) = .Item("cboAuth").ValueEx.ToString.Trim
                sInput(5) = Convert.ToInt32(.Item("tbPortNum").ValueEx)
                sInput(6) = .Item("tbReceipt").ValueEx.ToString.Trim
            End With

            'validate all
            If sInput(0) = "" Then
                SBO_Application.StatusBar.SetText("Email address of the sender is blank. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If
            If sInput(1) = "" Then
                SBO_Application.StatusBar.SetText("SMTP Outgoing Server is blank. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If
            If sInput(4) = "1" Then
                If sInput(2) = "" Then
                    SBO_Application.StatusBar.SetText("Username is empty. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Exit Sub
                End If
                If sInput(3) = "" Then
                    SBO_Application.StatusBar.SetText("Password is empty. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    Exit Sub
                End If
            End If
            If sInput(6) = "" Then
                SBO_Application.StatusBar.SetText("Email address of the recipient is blank. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Exit Sub
            End If

            'IRP HANA
            '=======================================================================================================
            Select Case oForm.DataSources.UserDataSources.Item("ckOutlook").ValueEx
                Case "Y"
                    Dim OutlookMessage As outlook.MailItem
                    Dim AppOutlook As New outlook.Application
                    Try
                        OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
                        Dim Recipients As outlook.Recipients = OutlookMessage.Recipients
                        Dim s() As String = sInput(6).Split(";")

                        For i As Integer = 0 To s.Length - 1
                            Recipients.Add(s(i).Trim)
                        Next

                        OutlookMessage.Subject = oCompany.CompanyName & " - Test Sending Email From MS Outlook - " & sToday
                        OutlookMessage.HTMLBody = "Test Connection - Successful"
                        OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                        OutlookMessage.Send()

                        SBO_Application.MessageBox("[MS Outlook] Test sending email is successful, please check the recipient's email to confirm.", 1, "OK")

                    Catch ex As Exception
                        SBO_Application.MessageBox("[MS Outlook] Sending Email Failed : " & ex.Message, 1, "OK")
                    Finally
                        OutlookMessage = Nothing
                        AppOutlook = Nothing
                    End Try

                    'Dim OutlookMessage As outlook.MailItem
                    'Dim AppOutlook As New outlook.Application
                    'Dim sOutput As String = ""
                    'Dim sOutput2 As String = ""
                    'Dim bImage1 As Boolean = False
                    'Dim bImage2 As Boolean = False
                    'Dim bImage3 As Boolean = False
                    'Dim bImage4 As Boolean = False
                    'Dim bImage5 As Boolean = False
                    'Dim bImage6 As Boolean = False
                    'Dim propertyAccessor As outlook.PropertyAccessor
                    'Dim image1 As outlook.Attachment
                    'Dim image2 As outlook.Attachment
                    'Dim image3 As outlook.Attachment
                    'Dim image4 As outlook.Attachment
                    'Dim image5 As outlook.Attachment
                    'Dim image6 As outlook.Attachment
                    'Dim attachments As outlook.Attachments = Nothing

                    'Try
                    '    sOutput = System.IO.File.ReadAllText("C:\Visual Studio Projects\IRP V905.091.2009 HANA\IRP V905.091.2009_GIT\Inecom SDK Reporting Package\bin\EmailBody.html")
                    '    bImage1 = CheckImage1(sOutput)
                    '    bImage2 = CheckImage2(sOutput)
                    '    bImage3 = CheckImage3(sOutput)
                    '    bImage4 = CheckImage4(sOutput)
                    '    bImage5 = CheckImage5(sOutput)
                    '    bImage6 = CheckImage6(sOutput)

                    '    OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
                    '    Dim Recipients As outlook.Recipients = OutlookMessage.Recipients
                    '    Dim s() As String = sInput(6).Split(";")

                    '    For i As Integer = 0 To s.Length - 1
                    '        Recipients.Add(s(i).Trim)
                    '    Next

                    '    If sOutput.Contains("{0}") Then
                    '        sOutput2 = sOutput.Replace("{0}", AsAtDate.ToString("dd/MM/yyyy"))
                    '    Else
                    '        sOutput2 = sOutput
                    '    End If

                    '    attachments = OutlookMessage.Attachments

                    '    If bImage1 And File.Exists() Then

                    '    End If
                    '    image1 = attachments.Add("C:\Visual Studio Projects\IRP V905.091.2009 HANA\IRP V905.091.2009_GIT\Inecom SDK Reporting Package\bin\image001.gif")
                    '    propertyAccessor = image1.PropertyAccessor
                    '    propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image1")

                    '    image2 = attachments.Add("C:\Visual Studio Projects\IRP V905.091.2009 HANA\IRP V905.091.2009_GIT\Inecom SDK Reporting Package\bin\image002.jpg")
                    '    propertyAccessor = image2.PropertyAccessor
                    '    propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image2")


                    '    image3 = attachments.Add(Directory.GetCurrentDirectory & "\image003.gif")
                    '    propertyAccessor = image3.PropertyAccessor
                    '    propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image3")

                    '    image4 = attachments.Add(Directory.GetCurrentDirectory & "\image004.gif")
                    '    propertyAccessor = image4.PropertyAccessor
                    '    propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image4")

                    '    image5 = attachments.Add(Directory.GetCurrentDirectory & "\image005.gif")
                    '    propertyAccessor = image5.PropertyAccessor
                    '    propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image5")

                    '    image6 = attachments.Add(Directory.GetCurrentDirectory & "\image006.gif")
                    '    propertyAccessor = image6.PropertyAccessor
                    '    propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image6")


                    '    OutlookMessage.Subject = oCompany.CompanyName & " - Test Sending Email From MS Outlook - " & sToday
                    '    OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                    '    OutlookMessage.HTMLBody = sOutput2
                    '    OutlookMessage.Send()

                    '    SBO_Application.MessageBox("Test sending email is successful, please check the recipient's email to confirm.", 1, "OK")

                    'Catch ex As Exception
                    '    SBO_Application.MessageBox("[MS Outlook] Sending Email Failed : " & ex.Message, 1, "OK")
                    'Finally
                    '    image1 = Nothing
                    '    image2 = Nothing
                    '    attachments = Nothing
                    '    OutlookMessage = Nothing
                    '    AppOutlook = Nothing
                    'End Try


                Case Else
                    Dim tmpMailFr As New System.Net.Mail.MailAddress(sInput(0))
                    Dim a As New System.Net.Mail.MailMessage()
                    Dim s() As String = sInput(6).Split(";")

                    a.From = tmpMailFr
                    a.Subject = oCompany.CompanyName & " - Test Sending Email From SMTP Connection - " & sToday
                    a.IsBodyHtml = True
                    a.Body = "Test Connection - Successful"


                    For i As Integer = 0 To s.Length - 1
                        a.To.Add(s(i).Trim)
                    Next


                    Dim c As New System.Net.Mail.SmtpClient(sInput(1))

                    If sInput(4) = "1" Then
                        Dim d As New System.Net.NetworkCredential(sInput(2), sInput(3))

                        If oForm.DataSources.UserDataSources.Item("ckOffice").ValueEx.ToString.Trim = "Y" Then
                            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                            c.EnableSsl = True
                        Else
                            Select Case sEnableSSL
                                Case "Y"
                                    System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                                    c.EnableSsl = True
                                Case "N"
                                    c.EnableSsl = False
                            End Select
                        End If

                        If sInput(5) > 0 Then
                            c.Port = sInput(5)
                        End If

                        c.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
                        c.UseDefaultCredentials = False
                        c.Credentials = d
                    End If

                    Try
                        c.Send(a)
                        SBO_Application.MessageBox("Test sending email is successful, please check the recipient's email to confirm.", 1, "OK")
                    Catch ex As Exception
                        SBO_Application.MessageBox("[SMTP] Sending Email Failed : " & ex.Message & " - " & ex.ToString, 1, "OK")
                    End Try

                    c = Nothing
                    a = Nothing

            End Select

            '=======================================================================================================
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[Test Connection] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Events Handler"
    Friend Function SBO_Application_ItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If pVal.Before_Action Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                BubbleEvent = ValidateBeforeSave()
                            End If
                        End If

                End Select
            Else 'After Action
                Select Case pVal.EventType

                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Select Case pVal.ItemUID
                            Case "flSMTP"
                                oForm.PaneLevel = 1
                                g_iPaneLvl = 1
                            Case "flHTML"
                                oForm.PaneLevel = 2
                                g_iPaneLvl = 2
                            Case "1"
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    If ValidateData() Then
                                        If Save() Then

                                            Dim iRow As Integer = 1
                                            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                            sQuery = "  SELECT ""Code"", IFNULL(""U_EmailAdd"",'') "
                                            sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCMCCEMAIL"" "
                                            sQuery &= " ORDER BY ""Code"" "

                                            oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecord.DoQuery(sQuery)
                                            If oRecord.RecordCount > 0 Then
                                                oRecord.MoveFirst()
                                                oMtrx.Clear()

                                                While Not oRecord.EoF
                                                    With oForm.DataSources.UserDataSources
                                                        .Item("cRow").ValueEx = iRow
                                                        .Item("cCode").ValueEx = oRecord.Fields.Item(0).Value.ToString.Trim
                                                        .Item("cEmail").ValueEx = oRecord.Fields.Item(1).Value.ToString.Trim
                                                    End With
                                                    iRow += 1
                                                    oMtrx.AddRow()
                                                    oRecord.MoveNext()
                                                End While
                                                oMtrx.AutoResizeColumns()
                                            End If

                                        Else
                                            BubbleEvent = False
                                        End If
                                    End If

                                End If
                            Case "btConnect"
                                Select Case pVal.FormMode
                                    Case SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        SBO_Application.StatusBar.SetText("Please update the fields into the database before testing the connection.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    Case SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        Connect()
                                End Select
                        End Select

                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        If String.Compare(pVal.ItemUID, "cboAuth", True) = 0 Then
                            If pVal.ItemChanged Then
                                If oForm.DataSources.UserDataSources.Item("cboAuth").ValueEx.Trim() = "0" Then
                                    oForm.Items.Item("txtUser").Enabled = False
                                    oForm.Items.Item("txtPass").Enabled = False
                                Else
                                    oForm.Items.Item("txtUser").Enabled = True
                                    oForm.Items.Item("txtPass").Enabled = True
                                End If ' End If "cboAuth" = 0
                            End If 'End If pval.ItemChanged
                        End If 'end if pval.ItemUID = "cboAuth"
                End Select 'End Select pval.EventType
            End If 'End If pVal.Before_Action

        Catch ex As Exception
            'SBO_Application.StatusBar.SetSystemMessage("[frmEmail].[ItemEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return BubbleEvent
    End Function
    Friend Function MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent) As Boolean
        Dim BubbleEvent As Boolean = True
        Try
            If (pVal.BeforeAction = True) Then
                Select Case pVal.MenuUID
                    Case MenuID.Delete_Row
                        Dim iCurrLine As Integer = 0
                        iCurrLine = oMtrx.GetNextSelectedRow

                        If iCurrLine <= 0 Or iCurrLine > oMtrx.VisualRowCount Then
                            BubbleEvent = False
                            SBO_Application.StatusBar.SetText("Select a valid line to delete.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                End Select
            Else
                Select Case pVal.MenuUID
                    Case MenuID.Delete_Row
                        oForm.Freeze(True)
                        For i As Integer = 1 To oMtrx.VisualRowCount Step 1
                            oMtrx.GetLineData(i)
                            oForm.DataSources.UserDataSources.Item("cRow").ValueEx = i
                            oMtrx.SetLineData(i)
                        Next
                        oForm.Freeze(False)

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                    Case MenuID.Add_Row
                        With oForm.DataSources.UserDataSources
                            .Item("cRow").ValueEx = oMtrx.VisualRowCount + 1
                            .Item("cCode").ValueEx = ""
                            .Item("cEmail").ValueEx = ""
                        End With
                        oMtrx.AddRow()
                        oMtrx.SelectRow(oMtrx.VisualRowCount, True, False)
                        oMtrx.Columns.Item("cEmail").Cells.Item(oMtrx.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                End Select
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            BubbleEvent = False
            SBO_Application.StatusBar.SetText("[MenuEvent] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        Return BubbleEvent
    End Function

#End Region

End Class