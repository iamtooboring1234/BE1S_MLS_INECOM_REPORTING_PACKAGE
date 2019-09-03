Imports System.IO
Imports EASendMail

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

            oForm.EnableMenu(MenuID.Add, False)
            oForm.EnableMenu(MenuID.Find, False)

            oForm.SupportedModes = -1
            oForm.Items.Item("tbReceipt").AffectsFormMode = False

            AddDataSource()
            SetDatasource()

            Try
                oForm.Items.Item("btConnect").Visible = True
            Catch ex As Exception

            End Try
            oForm.Visible = True
            oForm.Items.Item("tbReceipt").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
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
        End With

        oEdit = oForm.Items.Item("tbReceipt").Specific
        oEdit.DataBind.SetBound(True, "", "tbReceipt")
        oEdit = oForm.Items.Item("tbMailBody").Specific
        oEdit.DataBind.SetBound(True, "", "tbMailBody")
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

    End Sub
    Private Sub SetDatasource()
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = String.Empty

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
                    .Item("tbMailBody").ValueEx = oRecord.Fields.Item(5).Value
                    .Item("tbPortNum").ValueEx = oRecord.Fields.Item(6).Value
                    .Item("ckOffice").ValueEx = oRecord.Fields.Item(7).Value

                    If .Item("cboAuth").ValueEx.Trim() = "0" Then
                        oForm.Items.Item("txtUser").Enabled = False
                        oForm.Items.Item("txtPass").Enabled = False
                    Else
                        oForm.Items.Item("txtUser").Enabled = True
                        oForm.Items.Item("txtPass").Enabled = True
                    End If
                End With
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetSystemMessage("[Email - SetDataSources] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                        SBO_Application.StatusBar.SetText("Port Num is not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbPortNum"
                        Return False
                    End If
                End If

                If .Item("tbMailBody").ValueEx.Trim.Length > 0 Then
                    If Not System.IO.File.Exists(.Item("tbMailBody").ValueEx.Trim) Then
                        SBO_Application.StatusBar.SetText("File location is not valid. Please check.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.ActiveItem = "tbMailBody"
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
                    sInput(6) = .Item("tbMailBody").ValueEx
                    sInput(7) = .Item("ckOffice").ValueEx
                End With

                sQuery = String.Format(sQuery, sInput)
                oRecord = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecord.DoQuery(sQuery)
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
    Private Sub Connect()
        Try
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sRec As String = ""
            Dim sEnableSSL As String = ""
            Dim sEmailCc As String = GetEmailCCFromUDT()

            sRec = " SELECT TOP 1 IFNULL(""U_EnableSSL"",'N') FROM ""@NCM_EMAIL_CONFIG"" WHERE ""Code"" = 'SOA' "

            oRec.DoQuery(sRec)
            If oRec.RecordCount > 0 Then
                sEnableSSL = oRec.Fields.Item(0).Value
            End If

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
            'Select Case oForm.DataSources.UserDataSources.Item("ckOffice").ValueEx.ToString.Trim
            '    Case "Y"
            'Dim oMail As New SmtpMail("TryIt")
            'Dim oSmtp As New SmtpClient()
            'Dim s() As String = sInput(6).Split(";")

            '' Your hotmail/outlook email address
            'oMail.From = sInput(0)

            '' Set recipient email address, please change it to yours
            'For i As Integer = 0 To s.Length - 1
            '    oMail.To.Add((s(i).Trim))
            'Next

            'Dim Cc() As String
            'If sEmailCc.Trim.Length > 0 Then
            '    Cc = sEmailCc.Split(";")
            '    For i As Integer = 0 To Cc.Length - 1
            '        oMail.Cc.Add((Cc(i).Trim))
            '    Next
            'End If

            '' Set email subject
            'oMail.Subject = oCompany.CompanyName & " - Test Sending Email From SMTP Connection "

            '' Set email body
            'oMail.HtmlBody = "Test Connection - Successful"

            '' Hotmail/Outlook SMTP server address
            'Dim oServer As New SmtpServer(sInput(1))

            'oServer.User = sInput(2)
            'oServer.Password = sInput(3)

            '' use 587 port
            'oServer.Port = sInput(5)

            '' detect SSL/TLS connection automatically
            'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto

            'Try
            '    oSmtp.SendMail(oServer, oMail)
            'Catch ep As Exception
            '    SBO_Application.StatusBar.SetText("[Test Connection] : " & ep.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'End Try

            'Case "N"

            Dim tmpMailFr As New System.Net.Mail.MailAddress(sInput(0))
            Dim a As New System.Net.Mail.MailMessage()
            Dim s() As String = sInput(6).Split(";")
            Dim Cc() As String

            a.From = tmpMailFr
            a.Subject = oCompany.CompanyName & " - Test Sending Email From SMTP Connection "
            a.IsBodyHtml = True
            a.Body = "Test Connection - Successful"

            For i As Integer = 0 To s.Length - 1
                a.To.Add(s(i).Trim)
            Next

            If sEmailCc.Trim.Length > 0 Then
                Cc = sEmailCc.Split(";")
                For i As Integer = 0 To Cc.Length - 1
                    a.CC.Add((Cc(i).Trim))
                Next
            End If

            Dim c As New System.Net.Mail.SmtpClient(sInput(1))

            If sInput(4) = "1" Then
                Dim d As New System.Net.NetworkCredential(sInput(2), sInput(3))

                If oForm.DataSources.UserDataSources.Item("ckOffice").ValueEx.ToString.Trim = "Y" Then
                    c.EnableSsl = True
                Else
                    Select Case sEnableSSL
                        Case "Y"
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
                SBO_Application.MessageBox("Sending Email Failed : " & ex.Message, 1, "OK")
            End Try

            c = Nothing
            a = Nothing
            'End Select
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
                            Case "1"
                                If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    If Save() Then
                                        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        'Dim oBttn As SAPbouiCOM.Button = oForm.Items.Item("1").Specific
                                        'oBttn.Caption = "OK"
                                        'do nothing
                                    Else
                                        BubbleEvent = False
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
#End Region

End Class