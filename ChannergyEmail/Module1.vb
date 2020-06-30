Imports System.Data
Imports System.Data.Odbc
Imports System.IO
Imports System.Collections
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Timers
Imports System.Threading
Imports System.Net.Mail
Imports System.Net
Imports System.Windows.Forms
Imports BaiqiSoft.HtmlEditorControl
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports iTextSharp.text.html.simpleparser
Imports iTextSharp.tool.xml
Imports System.Drawing


Module Module1
    Public sqlArray(30) As String
    Public bolIsError As Boolean = False
    Public bolIsClientServer As Boolean = False
    Public stODBCString As String
    Public stPath As String = System.AppDomain.CurrentDomain.BaseDirectory()
    'Public stPath As String = "E:\Core\HeltonTools\Test"
    Public stEmailName As String
    Public stEmailSQL As String
    Public stdatafields(100) As String
    Public stInsertFields(100) As String
    Public stEmailField As String
    Public stGroupByField As String
    Public bolIsUpdateContactLog As Boolean
    Public bolIsActive As Boolean
    Public stFrequency As String
    Public stDays As String
    Public stCustNoField As String
    Public stCommandSQL As String
    Public bolIsPrompt As Boolean
    Public stDataType1 As String
    Public stDataType2 As String
    Public stPrompt1 As String
    Public stPrompt2 As String
    Public stValue1 As String
    Public stValue2 As String
    Public boIsAttachPDF As Boolean = False
    Public bolIsAttachCSV As Boolean = False
    Public stFieldarray() As String
    Public Sub ConvertHTML(oldfilename As String, newfilename As String)
        Dim pngfilename As String = Path.GetTempFileName()
        Dim res As String = "" ' = ok

        Try
            Using wb As System.Windows.Forms.WebBrowser = New System.Windows.Forms.WebBrowser
                wb.ScrollBarsEnabled = False
                wb.ScriptErrorsSuppressed = True
                wb.Navigate(oldfilename)
                While Not (wb.ReadyState = WebBrowserReadyState.Complete)
                    Application.DoEvents()
                End While

                wb.Width = wb.Document.Body.ScrollRectangle.Width
                wb.Width = PdfSharp.Drawing.XUnit.FromInch(8.0)
                wb.Width = wb.Width / 0.75
                wb.Height = wb.Document.Body.ScrollRectangle.Height
                'wb.Height = PdfSharp.Drawing.XUnit.FromInch(11)
                'wb.Height = wb.Height / 0.75

                If wb.Height > 3000 Then
                    wb.Height = 3000
                End If
                ' Get a Bitmap representation of the webpage as it's rendered in the WebBrowser control
                Dim b As Bitmap = New System.Drawing.Bitmap(wb.Width, wb.Height)
                Dim hr As Integer = b.HorizontalResolution
                Dim vr As Integer = b.VerticalResolution

                wb.DrawToBitmap(b, New System.Drawing.Rectangle(0, 0, wb.Width, wb.Height))
                wb.Dispose()
                If File.Exists(pngfilename) Then
                    File.Delete(pngfilename)
                End If
                b.Save(pngfilename, Imaging.ImageFormat.Png)
                b.Dispose()

                Using doc As PdfSharp.Pdf.PdfDocument = New PdfSharp.Pdf.PdfDocument
                    Dim page As PdfSharp.Pdf.PdfPage = New PdfSharp.Pdf.PdfPage()
                    page.Orientation = PdfSharp.PageOrientation.Portrait
                    page.Width = PdfSharp.Drawing.XUnit.FromInch(wb.Width / hr)
                    page.Height = PdfSharp.Drawing.XUnit.FromInch(wb.Height / vr)
                    page.Width = PdfSharp.Drawing.XUnit.FromInch(8.5)
                    page.Height = PdfSharp.Drawing.XUnit.FromInch(11)

                    doc.Pages.Add(page)

                    Dim xgr As PdfSharp.Drawing.XGraphics = PdfSharp.Drawing.XGraphics.FromPdfPage(page)
                    Dim img As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromFile(pngfilename)
                    'xgr.DrawImage(img, 0, 0)
                    xgr.DrawImage(img, 10, 10)
                    doc.Save(newfilename)
                    doc.Close()
                    img.Dispose()
                    xgr.Dispose()
                End Using

            End Using

        Catch ex As Exception
            res = "Error: " & ex.Message
            MessageBox.Show(res)
        Finally
            If File.Exists(pngfilename) Then
                File.Delete(pngfilename)
            End If

        End Try


    End Sub
    Sub SaveAsPDF(ByRef stHtmlFile As String, ByRef stPdfFile As String)
        Dim doc As New Document(iTextSharp.text.PageSize.LETTER, 50, 50, 50, 50)
        Dim sr As New StringReader(File.ReadAllText(stHtmlFile))
        Dim writer As PdfWriter = PdfWriter.GetInstance(doc, New FileStream(stPdfFile, FileMode.Create))
        Try
            doc.Open()
            XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, sr)
            doc.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub LoadEmail(ByRef stEmail As String)
        Dim stSQL As String = "SELECT * FROM EmailHtml WHERE EmailName='" + stEmail + "';"
        Dim con As New OdbcConnection(stODBCString)
        Dim cmd As OdbcCommandBuilder
        Dim daEmailList As New OdbcDataAdapter(stSQL, con)

        'Make sure all of the optional fields are not visible
        Form1.TxtPrompt1.Visible = False
        Form1.TxtPrompt2.Visible = False
        Form1.Date1.Visible = False
        Form1.Date2.Visible = False
        Form1.lblPrompt1.Visible = False
        Form1.lblPrompt2.Visible = False

        'Create a dataset and add the records to it
        Dim ds As New DataSet()
        daEmailList.Fill(ds, "EmailTemplate")

        For Each dr In ds.Tables(0).Rows
            stEmailSQL = dr.item("SQL").ToString
            If dr.item("IsPrompt").ToString = "True" Then
                bolIsPrompt = True
                stDataType1 = dr.item("DataType1").ToString
                stDataType2 = dr.item("DataType2").ToString
                stPrompt1 = dr.item("Prompt1").ToString
                stPrompt2 = dr.item("Prompt2").ToString

                'Unhide the fields based on the parameters
                If stDataType1 = "TEXT" Or stDataType1 = "INTEGER" Then
                    Form1.TxtPrompt1.Visible = True
                End If

                If stDataType1 = "DATE" Then
                    Form1.Date1.Visible = True
                    stValue1 = getDBIASMDate(DateTime.Now.ToString("MM/dd/yyyy"))

                End If

                Form1.lblPrompt1.Visible = True
                Form1.lblPrompt1.Text = stPrompt1

                If stPrompt2 <> "" Then
                    Form1.lblPrompt2.Visible = True
                    Form1.lblPrompt2.Text = stPrompt2
                    If stDataType2 = "TEXT" Or stDataType2 = "INTEGER" Then
                        Form1.TxtPrompt2.Visible = True
                    End If

                    If stDataType2 = "DATE" Then
                        Form1.Date2.Visible = True
                        stValue2 = getDBIASMDate(DateTime.Now.ToString("MM/dd/yyyy"))
                    End If
                End If
            Else
                bolIsPrompt = False
            End If

            'Code Added to set the boIsAttachPDF flag
            If dr.item("IsAttachmentAsPDF").ToString = "True" Then
                boIsAttachPDF = True
            End If

            If dr.item("IsAttachmentAsCSV").ToString = "True" Then

                bolIsAttachCSV = True
                LoadAttachmentFields()
            End If

            If dr.item("IsAttachmentAsCSV").ToString <> "True" And dr.item("IsAttachmentAsPDF").ToString <> "True" Then
                boIsAttachPDF = False
                bolIsAttachCSV = False
            End If
        Next
    End Sub
    Sub LoadAttachmentFields()
        Dim stSQL As String = "SELECT AttachmentSQL FROM EmailHtml WHERE EmailName='" + stEmailName + "';"
        Dim stFields As String


        SQLTextQuery("S", stSQL, stODBCString, 1)

        If sqlArray(0) <> "NoData" And sqlArray(0) <> "" Then
            stFields = sqlArray(0)
            stFieldarray = Strings.Split(stFields, "|")

        End If


    End Sub
    Function getDBIASMDate(ByRef dtDate As Date) As String
        Dim stDate As String = "'" + CStr(DatePart(DateInterval.Year, dtDate)) + "-" + CStr(DatePart(DateInterval.Month, dtDate)) + "-" + CStr(DatePart(DateInterval.Day, dtDate)) + "'"
        Return stDate

    End Function
 


    Sub LoadForm()
        Dim Count As Integer = 1
        Dim stSQL As String
  

        'Clear out existing combo boxes
        Form1.cboEmailName.Items.Clear()
        Form1.cboEmailName.Refresh()
        'Form1.cboSubject.Items.Clear()
        'Form1.cboSubject.Refresh()
        Form1.cboEmailName.Text = ""

        'Refresh the form
        Form1.Refresh()

        stSQL = "SELECT COUNT(EmailName) FROM EmailHtml WHERE IsPrompt=True;"
        SQLTextQuery("S", stSQL, stODBCString, 1)

        If sqlArray(0) <> "0" And sqlArray(0) <> "NoData" Then

            LoadCombo("EmailHTML", "EmailName", "EmailName", Form1.cboEmailName, "IsPrompt=True")


        Else
            MessageBox.Show("No emails have been set up.", "No Emails Set up", MessageBoxButtons.OK)
        End If
    End Sub
    Sub LoadCombo(ByRef stTableName As String, ByRef stFieldName As String, ByRef stOrderBy As String, ByRef cmbobox As ComboBox, Optional ByRef stWhere As String = "")
        Dim iRowCount As Integer
        Dim stSQL As String
        Dim iCounter As Integer = 1

        cmbobox.Items.Clear()

        'Get the number of items in the list
        If stWhere = "" Then
            stSQL = "SELECT COUNT('" + stFieldName + "') FROM " + stTableName + ";"
            SQLTextQuery("S", stSQL, stODBCString, 1)
            If sqlArray(0) <> "0" Or sqlArray(0) <> "" Then
                iRowCount = CInt(sqlArray(0))
                stSQL = "SELECT " + stFieldName + "," + stOrderBy + " INTO TempCombo FROM " + stTableName + " ORDER BY " + stOrderBy + ";"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                stSQL = "ALTER TABLE TempCombo ADD RowID AUTOINC;"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                Do While iCounter <= iRowCount
                    stSQL = "SELECT " + stFieldName + " FROM TempCombo WHERE RowID=" + CStr(iCounter) + ";"
                    SQLTextQuery("S", stSQL, stODBCString, 1)
                    cmbobox.Items.Add(sqlArray(0))

                    iCounter = iCounter + 1

                Loop
            End If
        Else
            stSQL = "SELECT COUNT('" + stFieldName + "') FROM " + stTableName + " WHERE " + stWhere + ";"
            SQLTextQuery("S", stSQL, stODBCString, 1)
            If sqlArray(0) <> "0" Or sqlArray(0) <> "" Then
                iRowCount = CInt(sqlArray(0))
                stSQL = "SELECT " + stFieldName + "," + stOrderBy + " INTO TempCombo FROM " + stTableName + " WHERE " + stWhere + " ORDER BY " + stOrderBy + ";"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                stSQL = "ALTER TABLE TempCombo ADD RowID AUTOINC;"
                SQLTextQuery("I", stSQL, stODBCString, 0)

                Do While iCounter <= iRowCount
                    stSQL = "SELECT " + stFieldName + " FROM TempCombo WHERE RowID=" + CStr(iCounter) + ";"
                    SQLTextQuery("S", stSQL, stODBCString, 1)
                    cmbobox.Items.Add(sqlArray(0))

                    iCounter = iCounter + 1

                Loop
            End If
        End If



    End Sub
    Sub UpdateLastSent()
        Dim stSQL As String = "SELECT EmailName,MAX(SentDt) AS LastSent INTO TempSent FROM EmailArchive WHERE Status='Sent' AND EmailName<>'' GROUP BY EmailName;"

        SQLTextQuery("U", stSQL, stODBCString, 0)

        stSQL = "CREATE INDEX IF NOT EXISTS idxEmailName ON TempSent(EmailName);"
        SQLTextQuery("U", stSQL, stODBCString, 0)

        stSQL = "UPDATE EmailHtml E SET LastSent=T.LastSent FROM EmailHtml E JOIN TempSent T ON E.EmailName=T.EmailName;"
        SQLTextQuery("U", stSQL, stODBCString, 0)


    End Sub
    Function FixHTML(ByRef stHTML) As String
        Dim stFixHTML As String
        stFixHTML = Replace(stHTML, "<p>&nbsp;</p>", "<br />")

        Return stFixHTML

    End Function
    Sub SendEmails()
        Dim stError As String
        Dim iEmailBatchNo As Integer
        Dim stSMTPCredentials(10) As String
        Dim stEmailHeader(10) As String
        Dim con As New OdbcConnection(stODBCString)
        Dim cmd As OdbcCommandBuilder
        Dim dt As DataTable
        Dim stEmailName As String
        Dim stTo As String
        Dim stSubject As String
        Dim ExportFile As System.IO.StreamWriter
        Dim stFilePath As String
        Dim stExportPath As String
        Dim stBodyHtml As String
        Dim stFileName As String




        'Set up the adapter for the compiled Emails to be sent
        Dim daEmailHtml As New OdbcDataAdapter("SELECT * FROM EmailBatchHtml", con)

        Dim ds As New DataSet()
        daEmailHtml.Fill(ds, "SendRecords")



        Array.Copy(GetSMTPCredentials(), stSMTPCredentials, stSMTPCredentials.Length)




        If stSMTPCredentials(0) <> "NoData" Then


            For Each dr In ds.Tables("SendRecords").Rows
                Try
                    'Reset the error flag
                    stError = "Success"
                    Dim HtmlEmail As New BaiqiSoft.HtmlEditorControl.MstHtmlEditor
                    HtmlEmail.LicenseKey = "Y3282Q20129WP195957P"

                    HtmlEmail.BodyHTML = dr.Item("MessageBody")
                    stBodyHtml = dr.Item("MessageBody")

                    'Get the email name
                    stEmailName = dr.Item("EmailName").ToString


                    stTo = dr.item("ToEmail").ToString
                    stSubject = dr.item("Subject").ToString

                    Array.Copy(GetEmailHeader(stEmailName), stEmailHeader, stEmailHeader.Length)


                    Dim mailMsg As New MailMessage
                    If dr.item("PdfTextBody").ToString = "" Then
                        mailMsg = HtmlEmail.GetMailMessage()
                    Else
                        mailMsg.IsBodyHtml = False
                        mailMsg.Body = dr.item("PdfTextBody").ToString
                    End If

                    mailMsg.HeadersEncoding = System.Text.Encoding.UTF8

                    mailMsg.To.Add(New MailAddress(dr.item("ToEmail").ToString, dr.item("ToEmail").ToString))
                    If stEmailHeader(3) <> "" Then
                        mailMsg.Bcc.Add(New MailAddress(stEmailHeader(3), stEmailHeader(3)))
                    End If

                    If stEmailHeader(4) <> "" Then
                        mailMsg.CC.Add(New MailAddress(stEmailHeader(4), stEmailHeader(4)))
                    End If

                    mailMsg.ReplyTo = New MailAddress(stEmailHeader(2), stEmailHeader(2))
                    mailMsg.From = New MailAddress(stEmailHeader(0), stEmailHeader(1))
                    mailMsg.Subject = dr.item("Subject").ToString

                    'If the attach email option is set create the attachment
                    If boIsAttachPDF = True Then
                        stFilePath = stPath + "Export.html"
                        stFileName = stSubject + "-" + Replace(DateString, "/", "") + "-" + Replace(TimeString, ":", "")
                        stExportPath = stPath + stFileName + ".pdf"

                        ExportFile = My.Computer.FileSystem.OpenTextFileWriter(stFilePath, False)
                        ExportFile.WriteLine(stBodyHtml)
                        ExportFile.Close()
                        ConvertHTML(stFilePath, stExportPath)
                        mailMsg.Attachments.Add(New Attachment(stExportPath))
                    End If

                    If bolIsAttachCSV = True Then
                        stFileName = stSubject + "-" + Replace(DateString, "/", "") + "-" + Replace(TimeString, ":", "")
                        stFilePath = stPath + stFileName + "Export.csv"
                        ExportCSV(stFilePath, stCommandSQL)
                        mailMsg.Attachments.Add(New Attachment(stFilePath))
                    End If

                    Dim mySmtpClient As SmtpClient = New SmtpClient(stSMTPCredentials(0), CInt(stSMTPCredentials(3)))
                    mySmtpClient.UseDefaultCredentials = False
                    mySmtpClient.Credentials = New NetworkCredential(stSMTPCredentials(1), stSMTPCredentials(2))

                    'Code added to update the ssl
                    If stSMTPCredentials(4) = "True" Then
                        mySmtpClient.EnableSsl = True
                    Else
                        mySmtpClient.EnableSsl = False
                    End If

                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                    mySmtpClient.Send(mailMsg)

                Catch ex As SmtpException
                    stError = ex.Message
                    bolIsError = True
                    AddEmailLog(stEmailName, "Error", True, stError)
                Catch ex As Exception
                    stError = ex.Message
                    bolIsError = True
                    AddEmailLog(stEmailName, "Error", True, stError)
                End Try
                iEmailBatchNo = CInt(dr.Item("EmailBatchNo").ToString)
                UpdateEmailBatch(iEmailBatchNo, stError)

            Next
            If bolIsError = False Then
                MessageBox.Show(stSubject, "Email Sent")
            Else
                MessageBox.Show(stError, "There were errors in sending " + stSubject)
            End If
        Else
            MessageBox.Show("No SMTP Email Client has been set up.")
        End If
        daEmailHtml.Dispose()


    End Sub
    Sub ExportCSV(ByRef stFilePath As String, ByRef stSQL As String)
        Dim con As New OdbcConnection(stODBCString)
        Dim cmd As New OdbcCommand(stSQL, con)
        Dim da As New OdbcDataAdapter(cmd)
        Dim ds As New DataSet()
        Dim dt As New DataTable
        Dim sb As StringBuilder = New StringBuilder

        Dim iColumnCount As Integer = stFieldarray.Count
        Dim i As Integer = 0
        Dim stFieldName As String

        If stFieldarray.Count > 0 Then
            For i = 0 To stFieldarray.Count - 1
                stFieldName = stFieldarray(i).ToString
                sb.Append("""" + stFieldName + """")

                If i = iColumnCount - 1 Then
                    sb.Append(" ")
                Else
                    sb.Append(",")
                End If

            Next
            sb.Append(vbCrLf)
        End If


        da.Fill(ds, "logn")
        iColumnCount = ds.Tables(0).Columns.Count


        'Add the data
        Dim row As DataRow

        For Each row In ds.Tables(0).Rows
            Dim ir As Integer = 0
            For i = 0 To stFieldarray.Count - 1
                stFieldName = stFieldarray(i).ToString
                For ir = 0 To iColumnCount - 1
                    If ds.Tables(0).Columns(ir).ToString = stFieldName Then
                        sb.Append("""" + row.Item(ir).ToString().Replace("""", """""") + """")

                        If ir = iColumnCount - 1 Then
                            sb.Append(" ")
                        Else
                            sb.Append(",")
                        End If
                    End If

                Next
            Next
            sb.Append(vbCrLf)
        Next

        File.WriteAllText(stFilePath, sb.ToString)



    End Sub
    Sub AddEmailLog(ByRef stEmailName As String, ByRef stStatus As String, ByRef IsError As Boolean, Optional ByRef stError As String = "")
        Dim stSQL As String = "INSERT INTO EmailSendLog(EmailName,Status,IsError,ErrorMessage) "
        Dim iSQL As String

        stError = Replace(stError, "'", "")
        If IsError = False Then
            iSQL = "VALUES('" + stEmailName + "','" + stStatus + "',False,'" + stError + "');"
        Else
            iSQL = "VALUES('" + stEmailName + "','" + stStatus + "',True,'" + stError + "');"
        End If

        stSQL = stSQL + iSQL

        SQLTextQuery("I", stSQL, stODBCString, 0)

    End Sub

    Sub UpdateEmailBatch(ByRef iEmailBatchNo As Integer, ByRef stResponse As String)
        Dim stSQL As String

        stResponse = Replace(stResponse, "'", "")

        If stResponse <> "Success" Then
            stSQL = "UPDATE EmailBatchHtml SET IsSuccess=False,Response='" + stResponse + "' WHERE EmailBatchNo=" + CStr(iEmailBatchNo) + ";"
        Else
            stSQL = "UPDATE EmailBatchHtml SET IsSuccess=True,Response='" + stResponse + "' WHERE EmailBatchNo=" + CStr(iEmailBatchNo) + ";"
        End If
        SQLTextQuery("U", stSQL, stODBCString, 0)

        'Get the customer number
        stSQL = "SELECT CustNo FROM EmailBatchHtml WHERE EmailBatchNo=" + CStr(iEmailBatchNo) + ";"
        SQLTextQuery("S", stSQL, stODBCString, 1)

        If sqlArray(0) <> "NoData" Then
            UpdateCust(CStr(iEmailBatchNo), sqlArray(0))
        End If

    End Sub
    Sub UpdateCust(ByRef stEmailBatchNo As String, ByRef stCustNo As String)
        Dim con As New OdbcConnection(stODBCString)
        Dim cmd As OdbcCommandBuilder
        Dim stEmailMessage As String
        Dim iCounter As Integer = 1
        Dim iCustNo As Integer
        Dim stSubject As String
        Dim stSQL As String = "SELECT MAX(LogNo) AS LogNo INTO NextCServ FROM CustServ;"
        Dim stExportPath As String = stPath + "Attachments\Customers\"
        Dim stFileName As String
        Dim stFilePath As String
        Dim ExportFile As System.IO.StreamWriter
        Dim stOrderNo As String

        'Make sure that the export path exsits
        If My.Computer.FileSystem.DirectoryExists(stExportPath) = False Then
            My.Computer.FileSystem.CreateDirectory(stExportPath)
        End If


        'Get a list of emails that need to be added to the contact log
        Dim daEmailHtml As New OdbcDataAdapter("SELECT * FROM EmailBatchHtml WHERE IsUpdateContactLog=True AND IsSuccess=True AND EmailBatchNO=" + stEmailBatchNo, con)

        'Data adapter for Customer Attach
        Dim daCustAttach As New OdbcDataAdapter("SELECT * FROM CustomerAttachments", con)


        'Add both to a dataset
        Dim ds As New DataSet
        daEmailHtml.Fill(ds, "EmailBatch")
        daCustAttach.Fill(ds, "CustAttach")



        For Each dr In ds.Tables("EmailBatch").Rows
            iCustNo = CInt(stCustNo)
            stEmailMessage = dr.item("MessageBody").ToString
            stSubject = dr.item("Subject").ToString
            stOrderNo = dr.item("OrderNo").ToString

            stFileName = "CustNo-" + stCustNo + "-" + stSubject + "-" + Replace(DateString, "/", "") + "-" + Replace(TimeString, ":", "") + ".html"
            stFilePath = stPath + "Export.html"
            stExportPath = stExportPath + stFileName + ".pdf"

            ExportFile = My.Computer.FileSystem.OpenTextFileWriter(stFilePath, False)
            ExportFile.WriteLine(stEmailMessage)
            ExportFile.Close()
            SaveAsPDF(stFilePath, stExportPath)

            'Add the records to the EmailBatch
            cmd = New OdbcCommandBuilder(daCustAttach)

            Dim tblCustAttach As DataTable
            tblCustAttach = ds.Tables("CustAttach")
            Dim newEmailRow As DataRow = tblCustAttach.NewRow()
            Try
                newEmailRow("CustNo") = iCustNo
                newEmailRow("Description") = stSubject
                newEmailRow("Type") = "File"
                newEmailRow("IsDefaultPath") = True
                newEmailRow("FileName") = stFileName + ".pdf"


                tblCustAttach.Rows.Add(newEmailRow)
                daCustAttach.Update(ds, "CustAttach")
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                bolIsError = True
                AddEmailLog(stSubject, "Error", True, ex.Message)
            End Try

            'Add the data to CustServ
            stSQL = "SELECT MAX(LogNo) AS LogNo INTO NextCServ FROM CustServ;"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            'If the LogNo is NULL make it a 1
            stSQL = "UPDATE NextCServ SET LogNo=IF(LogNo=NULL,0,LogNo);"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            If stOrderNo = "" Then
                stSQL = "INSERT INTO CustServ(CustNo,LogNo,Date,StatusFlag) "
                stSQL = stSQL + "SELECT " + stCustNo + ",LogNo+1,CURRENT_DATE,'" + stSubject + "' FROM NextCServ;"
                SQLTextQuery("U", stSQL, stODBCString, 0)
            Else
                stSQL = "INSERT INTO CustServ(CustNo,LogNo,Date,StatusFlag,LinkOrderNo) "
                stSQL = stSQL + "SELECT " + stCustNo + ",LogNo+1,CURRENT_DATE,'" + stSubject + "'," + stOrderNo + " FROM NextCServ;"
                SQLTextQuery("U", stSQL, stODBCString, 0)
            End If

            stSQL = "SELECT MAX(LogNo) AS LogNo INTO NextCServ FROM CustServ;"
            SQLTextQuery("U", stSQL, stODBCString, 0)

            iCounter = iCounter + 1
        Next

        daEmailHtml.Dispose()
        con.Close()

    End Sub
    Sub LoadEmailHtmlBatch(ByRef stEmailSQL As String)
        Dim con As New OdbcConnection(stODBCString)
        Dim cmd As OdbcCommandBuilder


        'Get the Records for the emails to be send
        Dim daEmailList As New OdbcDataAdapter(stEmailSQL, con)

        'Set up the adapter for the compiled Emails to be sent
        Dim daEmailHtml As New OdbcDataAdapter("SELECT * FROM EmailBatchHtml", con)




        'Create a dataset and add the records to it
        Dim ds As New DataSet()
        daEmailList.Fill(ds, "SendRecords")
        daEmailHtml.Fill(ds, "EmailBatch")



        Dim dt As New DataTable
        Dim dtEmailBatch As New DataTable

        Dim dc As DataColumn
        Dim dr As DataRow
        Dim stSQL As String
        Dim stFromEmail As String
        Dim stFromName As String
        Dim stReplyEmail As String
        Dim stSubject As String
        Dim stBcc As String
        Dim stCC As String
        Dim stGroupByField As String
        Dim stEmailField As String
        Dim stSendTo As String = ""
        Dim stGroupBy As String = ""
        Dim iRowCount As Integer = 1
        Dim stColumn As String
        Dim stValue As String
        Dim stHeaderTemplate As String = StripHtml(GetTemplate(stEmailName, "Header"))
        Dim stDetailTemplate As String = StripHtml(GetTemplate(stEmailName, "Detail"))
        Dim stFooterTemplate As String = StripHtml(GetTemplate(stEmailName, "Footer"))
        Dim stEmailTemplate As String = File.ReadAllText(stPath + "EmailTemplate.html")
        Dim stHeader As String = ""
        Dim stDetail() As String
        Dim stFooter As String = ""
        Dim stDetailFields() As String
        Dim stTHeader As String
        'Dim stTHeader As String = "<table><tbody>"
        Dim stTFooter As String = "</tbody></table>"
        Dim stDetailTable As String = ""
        Dim iNumRecords As Integer
        Dim stEmailBody As String
        Dim stUpdateContactLog As String
        Dim stCustNoField As String
        Dim stOrderNoField As String
        Dim stOrderNo As String
        Dim stCustNo As String
        Dim bolDetail As Boolean
        Dim stDataType As String
        Dim dValue As Double
        Dim stSubjectTemplate As String
        Dim stPdfText As String

        'Add code to handle emails that do not have detail records
        If stDetailTemplate <> "" Then
            stDetail = GetDetails(stDetailTemplate)
            stDetailFields = GetDetails(stDetailTemplate)
            stTHeader = "<table " + GetDetailTag(stDetailTemplate) + "><tbody>"
            bolDetail = True
        Else
            bolDetail = False
        End If

        'Empty out the records from the EmailBachHtml
        SQLTextQuery("D", "DELETE FROM EmailBatchHtml;", stODBCString, 0)

        'Get the information from the EmailHtml table.
        stSQL = "SELECT FromEmailAddress,FromName,ReplyEmailAddress,Subject,BlindCarbonCopy,CarbonCopy,GroupByField,EmailAddressField,IsUpdateContactLog,CustNoField,OrderNoField,PdfTextBody FROM EmailHtml WHERE EmailName='" + stEmailName + "';"
        SQLTextQuery("S", stSQL, stODBCString, 12)

        stFromEmail = sqlArray(0)
        stFromName = sqlArray(1)
        stReplyEmail = sqlArray(2)
        stSubjectTemplate = sqlArray(3)
        stBcc = sqlArray(4)
        stCC = sqlArray(5)
        stGroupByField = sqlArray(6)
        stEmailField = sqlArray(7)
        stUpdateContactLog = sqlArray(8)
        stCustNoField = sqlArray(9)
        stOrderNoField = sqlArray(10)
        stPdfText = sqlArray(11)

        'Define the data tables
        dt = ds.Tables("SendRecords")
        dtEmailBatch = ds.Tables("EmailBatch")

        'Get the number of records to be processed
        iNumRecords = dt.Rows.Count


        For Each dr In ds.Tables("SendRecords").Rows


            'If this is the first record in the list then start here
            If iRowCount = 1 Then
                'Get the field that the email notification is grouped by
                stGroupBy = dr.Item(stGroupByField).ToString
                stSendTo = dr.Item(stEmailField).ToString
                If stCustNoField <> "" Then
                    stCustNo = dr.Item(stCustNoField).ToString
                End If

                If stOrderNoField <> "" Then
                    stOrderNo = dr.Item(stOrderNoField).ToString
                End If

                stHeader = stHeaderTemplate
                stDetailTable = stTHeader
                stFooter = stFooterTemplate
                stSubject = stSubjectTemplate

                'Update the header and footer information
                For Each dc In ds.Tables(0).Columns
                    stColumn = "[" + dc.ColumnName + "]"
                    stDataType = dc.DataType.ToString
                    stValue = dr.Item(dc.ColumnName).ToString
                    If stValue <> "" Then
                        If stDataType = "System.DateTime" Then
                            stValue = FormatDateTime(stValue, DateFormat.ShortDate)
                        ElseIf stDataType = "System.Double" Then
                            dValue = CDbl(stValue)
                            If Right(dc.ColumnName, 3) = "Amt" Or Right(dc.ColumnName, 5) = "Price" Or Right(dc.ColumnName, 4) = "Cost" Then 'this is a currency
                                stValue = FormatCurrency(dValue, 2, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.True)
                            End If
                        End If
                    End If
                    stHeader = Replace(stHeader, stColumn, stValue)
                    stFooter = Replace(stFooter, stColumn, stValue)
                    stSubject = Replace(stSubject, stColumn, stValue)
                Next


                'If there are detail records add them

                If bolDetail = True Then
                    Array.Copy(stDetailFields, stDetail, stDetail.Length)

                    'Get first detail record

                    For Each dc In ds.Tables(0).Columns
                        stColumn = "[" + dc.ColumnName + "]"
                        stValue = dr.Item(dc.ColumnName).ToString
                        stDataType = dc.DataType.ToString

                        If stValue <> "" Then
                            If stDataType = "System.DateTime" Then
                                stValue = FormatDateTime(stValue, DateFormat.ShortDate)
                            ElseIf stDataType = "System.Double" Then
                                dValue = CDbl(stValue)
                                If Right(dc.ColumnName, 3) = "Amt" Or Right(dc.ColumnName, 5) = "Price" Or Right(dc.ColumnName, 4) = "Cost" Then 'this is a currency
                                    stValue = FormatCurrency(dValue, 2, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.True)
                                End If
                            End If
                        End If

                        For i As Integer = 0 To 9
                            stDetail(i) = Replace(stDetail(i), stColumn, stValue)
                        Next

                    Next

                    For i As Integer = 0 To 9
                        If stDetail(i) <> "" Then
                            stDetailTable = stDetailTable + stDetail(i)
                        End If
                    Next
                End If


            ElseIf stGroupBy <> dr.Item(stGroupByField).ToString Then 'This is a new email
                stEmailBody = stEmailTemplate
                stGroupBy = dr.Item(stGroupByField).ToString


                'Close the details table
                If bolDetail = True Then
                    stDetailTable = stDetailTable + stTFooter
                End If

                'Update the email template file
                stEmailBody = Replace(stEmailBody, "[Header]", stHeader)
                stEmailBody = Replace(stEmailBody, "[Detail]", stDetailTable)
                stEmailBody = Replace(stEmailBody, "[Footer]", stFooter)

                'Add the records to the EmailBatch
                cmd = New OdbcCommandBuilder(daEmailHtml)

                Dim tblEmailBatch As DataTable
                tblEmailBatch = ds.Tables("EmailBatch")
                Dim newEmailRow As DataRow = tblEmailBatch.NewRow()
                Try
                    newEmailRow("EmailName") = stEmailName
                    newEmailRow("Subject") = stSubject
                    newEmailRow("ToEmail") = stSendTo
                    newEmailRow("MessageBody") = FixHTML(stEmailBody)

                    If stCustNoField <> "" Then
                        newEmailRow("CustNo") = stCustNo
                    End If

                    If stOrderNoField <> "" Then
                        newEmailRow("OrderNo") = stOrderNo
                    End If

                    newEmailRow("IsUpdateContactLog") = stUpdateContactLog
                    newEmailRow("PdfTextBody") = stPdfText


                    tblEmailBatch.Rows.Add(newEmailRow)
                    daEmailHtml.Update(ds, "EmailBatch")
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    bolIsError = True
                    AddEmailLog(stEmailName, "Error", True, ex.Message)
                End Try


                'Get the information for the new email
                stSendTo = dr.Item(stEmailField).ToString
                If stCustNoField <> "" Then
                    stCustNo = dr.Item(stCustNoField).ToString
                End If

                If stOrderNoField <> "" Then
                    stOrderNo = dr.Item(stOrderNoField).ToString
                End If

                stHeader = stHeaderTemplate
                stFooter = stFooterTemplate
                stGroupBy = dr.Item(stGroupByField).ToString
                stSubject = stSubjectTemplate

                'If there are detail records then copy the Array
                If bolDetail = True Then
                    Array.Copy(stDetailFields, stDetail, stDetail.Length)
                    stDetailTable = stTHeader
                End If



                For Each dc In ds.Tables(0).Columns
                    stColumn = "[" + dc.ColumnName + "]"
                    stValue = dr.Item(dc.ColumnName).ToString


                    stDataType = dc.DataType.ToString

                    If stValue <> "" Then
                        If stDataType = "System.DateTime" Then
                            stValue = FormatDateTime(stValue, DateFormat.ShortDate)
                        ElseIf stDataType = "System.Double" Then
                            dValue = CDbl(stValue)
                            If Right(dc.ColumnName, 3) = "Amt" Or Right(dc.ColumnName, 5) = "Price" Or Right(dc.ColumnName, 4) = "Cost" Then 'this is a currency
                                stValue = FormatCurrency(dValue, 2, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.True)
                            End If
                        End If
                    End If

                    stHeader = Replace(stHeader, stColumn, stValue)
                    stFooter = Replace(stFooter, stColumn, stValue)
                    stSubject = Replace(stSubject, stColumn, stValue)

                    If bolDetail = True Then
                        For i As Integer = 0 To 9
                            stDetail(i) = Replace(stDetail(i), stColumn, stValue)
                        Next
                    End If


                Next

                If bolDetail = True Then


                    'Add the detail record
                    For i As Integer = 0 To 9
                        If stDetail(i) <> "" Then
                            stDetailTable = stDetailTable + stDetail(i)
                        End If
                    Next
                End If


            ElseIf bolDetail = True And stGroupBy = dr.Item(stGroupByField).ToString And iRowCount > 1 Then 'This is a detail record
                Array.Copy(stDetailFields, stDetail, stDetail.Length)
                For Each dc In ds.Tables(0).Columns
                    stColumn = "[" + dc.ColumnName + "]"
                    stValue = dr.Item(dc.ColumnName).ToString
                    'stHeader = Replace(stHeader, stColumn, stValue)
                    'stFooter = Replace(stFooter, stColumn, stValue)
                    stDataType = dc.DataType.ToString

                    If stValue <> "" Then
                        If stDataType = "System.DateTime" Then
                            stValue = FormatDateTime(stValue, DateFormat.ShortDate)
                        ElseIf stDataType = "System.Double" Then
                            dValue = CDbl(stValue)
                            If Right(dc.ColumnName, 3) = "Amt" Or Right(dc.ColumnName, 5) = "Price" Or Right(dc.ColumnName, 4) = "Cost" Then 'this is a currency
                                stValue = FormatCurrency(dValue, 2, Microsoft.VisualBasic.TriState.True, Microsoft.VisualBasic.TriState.True)
                            End If
                        End If
                    End If


                    For i As Integer = 0 To 9
                        stDetail(i) = Replace(stDetail(i), stColumn, stValue)
                    Next

                Next
                'Add the detail record
                For i As Integer = 0 To 9
                    If stDetail(i) <> "" Then
                        stDetailTable = stDetailTable + stDetail(i)
                    End If
                Next
            End If

            'If this is the last record in the batch make sure that the data are added to the EmailBatch table
            If iRowCount = iNumRecords Then
                'Update the stEmailBody with the stEmailTemplate
                stEmailBody = stEmailTemplate

                'Close the details table
                stDetailTable = stDetailTable + stTFooter

                'Update the email template file
                stEmailBody = Replace(stEmailBody, "[Header]", stHeader)
                stEmailBody = Replace(stEmailBody, "[Detail]", stDetailTable)
                stEmailBody = Replace(stEmailBody, "[Footer]", stFooter)

                'Add the records to the EmailBatch
                cmd = New OdbcCommandBuilder(daEmailHtml)

                Dim tblEmailBatch As DataTable
                tblEmailBatch = ds.Tables("EmailBatch")
                Dim newEmailRow As DataRow = tblEmailBatch.NewRow()
                Try
                    newEmailRow("EmailName") = stEmailName
                    newEmailRow("Subject") = stSubject
                    newEmailRow("ToEmail") = stSendTo
                    newEmailRow("MessageBody") = FixHTML(stEmailBody)

                    If stCustNoField <> "" Then
                        newEmailRow("CustNo") = stCustNo
                    End If

                    If stOrderNoField <> "" Then
                        newEmailRow("OrderNo") = stOrderNo
                    End If

                    newEmailRow("IsUpdateContactLog") = stUpdateContactLog
                    newEmailRow("PdfTextBody") = stPdfText

                    tblEmailBatch.Rows.Add(newEmailRow)
                    daEmailHtml.Update(ds, "EmailBatch")
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    bolIsError = True
                End Try
            End If
            iRowCount = iRowCount + 1
        Next
    End Sub
    Function GetDetailTag(ByRef stDetail As String) As String
        Dim stFind As String = "(?<=<table)(.*?)(?=>)"
        Dim doregex2 As MatchCollection = Regex.Matches(stDetail, stFind)
        Dim stMatch As String = ""

        For Each match As Match In doregex2
            stMatch = stMatch + match.ToString
        Next

        Return stMatch
    End Function
    Sub AppendArchive()
        Dim stSQL As String = "INSERT INTO EmailArchive(EmailBatchNo,Status,ErrorMessage,SentDt,ToEmailAddress,EmailBody,CustNo,OrderNo,Attachment1,EmailName) "

        stSQL = stSQL + "SELECT EmailBatchNo,IF(IsSuccess=True,'Sent','Error'),Response,CURRENT_DATE,ToEmail,MessageBody,E.CustNo,E.OrderNo,C.FileName,E.EmailName "
        stSQL = stSQL + "FROM EmailBatchHtml E LEFT OUTER JOIN CustomerAttachments C ON E.CustNo=C.CustNo;"

        SQLTextQuery("U", stSQL, stODBCString, 0)




    End Sub
    Sub SQLTextQuery(ByVal QueryType As String, ByVal CommandText As String, ByVal stODBC As String, Optional ByVal Columns As Integer = 0)
        'Dim DBCString As String = "Driver={C:\dbisam\odbc\std\ver4\lib\dbodbc\dbodbc.dll};connectiontype=Local;remoteipaddress=127.0.0.1;RemotePort=12005;remotereadahead=50;catalogname=" + stODBCString + ";readonly=False;lockretrycount=15;lockwaittime=100;forcebufferflush=False;strictchangedetection=False;"

        'Dim DBC As New System.Data.Odbc.OdbcConnection
        'DBC.ConnectionString = DBCString

        Dim DBC As New OdbcConnection(stODBC)
        'Dim DBCString As String = "Dsn=" & stODBC & ";"

        'Dim DBC As New OdbcConnection(DBCString)

        If QueryType = "S" And Columns > 0 Then
            Try
                Dim SQL1 As New OdbcCommand
                SQL1.Connection = DBC
                SQL1.CommandType = CommandType.Text
                SQL1.CommandText = CommandText
                DBC.Open()

                Dim DataRow As OdbcDataReader
                DataRow = SQL1.ExecuteReader()
                DataRow.Read()
                If DataRow.HasRows Then
                    Dim Counter As Integer = 0
                    While Counter < Columns
                        sqlArray(Counter) = DataRow(Counter).ToString
                        Counter = Counter + 1
                    End While
                Else
                    sqlArray(0) = "NoData"
                End If
                DataRow.Close()
                DBC.Close()
                SQL1.Dispose()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                bolIsError = True
            End Try
        End If
        If QueryType = "U" Or QueryType = "I" Or QueryType = "D" Then
            Try
                DBC.Open()
                Dim SQL2 As New OdbcCommand
                SQL2.Connection = DBC
                SQL2.CommandType = CommandType.Text
                SQL2.CommandTimeout = 60
                SQL2.CommandText = CommandText
                SQL2.ExecuteScalar()
                DBC.Close()
                SQL2.Dispose()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                bolIsError = True
            End Try
        End If
    End Sub
    Function GetEmailHeader(ByRef stEmailName As String) As String()
        Dim stSQL As String = "SELECT FromEmailAddress,FromName,ReplyEmailAddress,BlindCarbonCopy,CarbonCopy FROM EmailHtml WHERE EmailName='" + stEmailName + "';"
        SQLTextQuery("S", stSQL, stODBCString, 5)

        Return sqlArray
    End Function

    Function GetSMTPCredentials() As String()
        Dim stSQL As String = "SELECT MailHost,AccountName,Password,Port,IsSSL,FromEmailAddress,FromName,ReplyEmailAddress,BlindCarbonCopy,CarbonCopy FROM EmailServerSettings;"
        SQLTextQuery("S", stSQL, stODBCString, 10)

        Return sqlArray

    End Function
    Function GetTemplate(ByRef stEmail As String, ByRef stType As String) As String
        Dim stSQL As String = "SELECT * FROM EmailHtml WHERE EmailName='" + stEmail + "';"
        Dim con As New OdbcConnection(stODBCString)
        Dim cmd As OdbcCommandBuilder
        Dim daEmailList As New OdbcDataAdapter(stSQL, con)

        Dim ds As New DataSet()
        daEmailList.Fill(ds, "EmailTemplate")

        For Each dr In ds.Tables(0).Rows
            If stType = "Header" Then
                Return dr.item("MessageHeader").ToString
            ElseIf stType = "Detail" Then
                Return dr.item("MessageDetail").ToString
            Else
                Return dr.item("MessageFooter").ToString
            End If
        Next

    End Function
    Function StripHtml(ByRef stHtml As String) As String
        stHtml = Replace(stHtml, "<html>", vbNullString)
        stHtml = Replace(stHtml, "</html>", vbNullString)
        stHtml = Replace(stHtml, "    <head>", vbNullString)
        stHtml = Replace(stHtml, "    </head>", vbNullString)
        stHtml = Replace(stHtml, "    <body>", vbNullString)
        stHtml = Replace(stHtml, "    </body>", vbNullString)

        Return stHtml
    End Function
    Function GetDetails(ByRef stDetail As String) As String()

        Dim HTMLDoc As New WebBrowser
        Dim Elems As HtmlElementCollection
        Dim counter As Integer = 0
        Dim stDetailArray(10) As String
        Dim stArrayValue As String


        HTMLDoc.Navigate(String.Empty)
        HTMLDoc.Document.Write(stDetail)


        Elems = HTMLDoc.Document.GetElementsByTagName("TD")

        For Each elem As HtmlElement In Elems
            stDetailArray(counter) = elem.OuterHtml.ToString
            counter = counter + 1
        Next

        Dim stReg As String = "<td(.*)/td>"


        'Dim mc As MatchCollection = Regex.Matches(stDetail, stReg)
        'Dim m As Match
        'Dim stArrayValue As String

        'For Each m In mc
        '    stDetailArray(counter) = m.ToString
        '    counter = counter + 1
        'Next m

        'Add the beginning and ending tags
        stArrayValue = "<tr>" + stDetailArray(0)
        stDetailArray(0) = stArrayValue

        stArrayValue = stDetailArray(counter - 1) + "</tr>"
        stDetailArray(counter - 1) = stArrayValue


        Return stDetailArray
    End Function
    Function GetDataPath(ByRef stExePath As String) As String
        Dim stIniFilePath As String
        Dim stIniFile = New IniFile()
        Dim stDatapath As String
        Dim virtualFilePath As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VirtualStore\Program Files (x86)\")
        Dim stVDataPath As String
        Dim stProgram As String

        'Get the version path and append to the virtualFilePath
        If Right(stExePath, 1) = "\" Then
            stProgram = Left(stExePath, Len(stExePath) - 1)
        Else
            stProgram = stExePath
        End If
        stProgram = Right(stProgram, Len(stProgram) - InStrRev(stProgram, "\"))
        virtualFilePath = virtualFilePath + stProgram


        'Check to see if the AppData folder contains the ini files
        If My.Computer.FileSystem.DirectoryExists(virtualFilePath) = True Then
            stVDataPath = virtualFilePath
        Else
            stVDataPath = stExePath
        End If


        'Fix path if it does not have a trailing \
        If Right(stVDataPath, 1) <> "\" Then
            stVDataPath = stVDataPath + "\"
        End If


        'Check to see if they are using Channergy or Mailware
        If My.Computer.FileSystem.FileExists(stVDataPath + "MAILWARE.INI") Then
            stIniFilePath = stVDataPath + "MAILWARE.INI"
        ElseIf My.Computer.FileSystem.FileExists(stVDataPath + "CHANNERGY.INI") Then
            stIniFilePath = stVDataPath + "CHANNERGY.INI"
        Else
            stIniFilePath = ""
        End If

        If stIniFilePath <> "" Then
            stIniFile.Load(stIniFilePath)
            stDatapath = stIniFile.GetKeyValue("Network", "NetworkDataDirectory")
            'Fix path if it does not have a trailing \
            If Right(stDatapath, 1) <> "\" Then
                stDatapath = stDatapath + "\"
            End If
        Else
            stDatapath = ""
        End If
        Return stDatapath

    End Function
    Function GetDSN(ByRef stDataPath As String) As String
        Dim stDSN As String
        Dim stIniFile = New IniFile()
        Dim stIniFilePath As String = stDataPath + "clientserver.ini"
        Dim stIpAddress As String
        Dim stCatalog As String
        Dim builder As New OdbcConnectionStringBuilder()

        builder.Driver = "DBISAM 4 ODBC Driver"

        'Check to see if the clientserver.ini file is in the passed path
        If My.Computer.FileSystem.FileExists(stIniFilePath) = True Then
            stIniFile.Load(stIniFilePath)
            stIpAddress = stIniFile.GetKeyValue("Settings", "IPAddress")
            stCatalog = stIniFile.GetKeyValue("Settings", "RemoteDatabase")
            builder.Add("UID", "Admin")
            builder.Add("PWD", "DBAdmin")
            builder.Add("ConnectionType", "Remote")
            builder.Add("RemoteIPAddress", stIpAddress)
            builder.Add("CatalogName", stCatalog)
            stDSN = builder.ConnectionString
            bolIsClientServer = True
        ElseIf My.Computer.FileSystem.FileExists(stDataPath + "Version.dat") = True Then
            builder.Add("ConnectionType", "Local")
            builder.Add("CatalogName", stDataPath)
            stDSN = builder.ConnectionString
        Else
            stDSN = ""
        End If

        Return stDSN
    End Function
    ' IniFile class used to read and write ini files by loading the file into memory
    Public Class IniFile
        ' List of IniSection objects keeps track of all the sections in the INI file
        Private m_sections As Hashtable

        ' Public constructor
        Public Sub New()
            m_sections = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
        End Sub

        ' Loads the Reads the data in the ini file into the IniFile object
        Public Sub Load(ByVal sFileName As String, Optional ByVal bMerge As Boolean = False)
            If Not bMerge Then
                RemoveAllSections()
            End If
            '  Clear the object... 
            Dim tempsection As IniSection = Nothing
            Dim oReader As New StreamReader(sFileName)
            Dim regexcomment As New Regex("^([\s]*#.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            ' Broken but left for history
            'Dim regexsection As New Regex("\[[\s]*([^\[\s].*[^\s\]])[\s]*\]", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexsection As New Regex("^[\s]*\[[\s]*([^\[\s].*[^\s\]])[\s]*\][\s]*$", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            Dim regexkey As New Regex("^\s*([^=\s]*)[^=]*=(.*)", (RegexOptions.Singleline Or RegexOptions.IgnoreCase))
            While Not oReader.EndOfStream
                Dim line As String = oReader.ReadLine()
                If line <> String.Empty Then
                    Dim m As Match = Nothing
                    If regexcomment.Match(line).Success Then
                        m = regexcomment.Match(line)
                        Trace.WriteLine(String.Format("Skipping Comment: {0}", m.Groups(0).Value))
                    ElseIf regexsection.Match(line).Success Then
                        m = regexsection.Match(line)
                        Trace.WriteLine(String.Format("Adding section [{0}]", m.Groups(1).Value))
                        tempsection = AddSection(m.Groups(1).Value)
                    ElseIf regexkey.Match(line).Success AndAlso tempsection IsNot Nothing Then
                        m = regexkey.Match(line)
                        Trace.WriteLine(String.Format("Adding Key [{0}]=[{1}]", m.Groups(1).Value, m.Groups(2).Value))
                        tempsection.AddKey(m.Groups(1).Value).Value = m.Groups(2).Value
                    ElseIf tempsection IsNot Nothing Then
                        '  Handle Key without value
                        Trace.WriteLine(String.Format("Adding Key [{0}]", line))
                        tempsection.AddKey(line)
                    Else
                        '  This should not occur unless the tempsection is not created yet...
                        Trace.WriteLine(String.Format("Skipping unknown type of data: {0}", line))
                    End If
                End If
            End While
            oReader.Close()
        End Sub
        ' Used to save the data back to the file or your choice
        Public Sub ListSections(ByVal sSection As String)
            Dim Counter As Integer = 0

            For Each s As IniSection In Sections
                If s.Name = sSection Then
                    For Each k As IniSection.IniKey In s.Keys
                        If k.Value <> String.Empty Then
                            'stCompanyList(Counter, 0) = k.Name
                            'stCompanyList(Counter, 1) = k.Value
                            'stCompanyList(Counter, 2) = CStr(Counter + 1)
                            Counter = Counter + 1
                        End If
                    Next
                End If
            Next

        End Sub
        ' Used to save the data back to the file or your choice
        Public Sub Save(ByVal sFileName As String)
            Dim oWriter As New StreamWriter(sFileName, False)
            For Each s As IniSection In Sections
                Trace.WriteLine(String.Format("Writing Section: [{0}]", s.Name))
                oWriter.WriteLine(String.Format("[{0}]", s.Name))
                For Each k As IniSection.IniKey In s.Keys
                    If k.Value <> String.Empty Then
                        Trace.WriteLine(String.Format("Writing Key: {0}={1}", k.Name, k.Value))
                        oWriter.WriteLine(String.Format("{0}={1}", k.Name, k.Value))
                    Else
                        Trace.WriteLine(String.Format("Writing Key: {0}", k.Name))
                        oWriter.WriteLine(String.Format("{0}", k.Name))
                    End If
                Next
            Next
            oWriter.Close()
        End Sub

        ' Gets all the sections
        Public ReadOnly Property Sections() As System.Collections.ICollection
            Get
                Return m_sections.Values
            End Get
        End Property

        ' Adds a section to the IniFile object, returns a IniSection object to the new or existing object
        Public Function AddSection(ByVal sSection As String) As IniSection
            Dim s As IniSection = Nothing
            sSection = sSection.Trim()
            ' Trim spaces
            If m_sections.ContainsKey(sSection) Then
                s = DirectCast(m_sections(sSection), IniSection)
            Else
                s = New IniSection(Me, sSection)
                m_sections(sSection) = s
            End If
            Return s
        End Function

        ' Removes a section by its name sSection, returns trus on success
        Public Function RemoveSection(ByVal sSection As String) As Boolean
            sSection = sSection.Trim()
            Return RemoveSection(GetSection(sSection))
        End Function

        ' Removes section by object, returns trus on success
        Public Function RemoveSection(ByVal Section As IniSection) As Boolean
            If Section IsNot Nothing Then
                Try
                    m_sections.Remove(Section.Name)
                    Return True
                Catch ex As Exception
                    Trace.WriteLine(ex.Message)
                End Try
            End If
            Return False
        End Function

        '  Removes all existing sections, returns trus on success
        Public Function RemoveAllSections() As Boolean
            m_sections.Clear()
            Return (m_sections.Count = 0)
        End Function

        ' Returns an IniSection to the section by name, NULL if it was not found
        Public Function GetSection(ByVal sSection As String) As IniSection
            sSection = sSection.Trim()
            ' Trim spaces
            If m_sections.ContainsKey(sSection) Then
                Return DirectCast(m_sections(sSection), IniSection)
            End If
            Return Nothing
        End Function

        '  Returns a KeyValue in a certain section
        Public Function GetKeyValue(ByVal sSection As String, ByVal sKey As String) As String
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.Value
                End If
            End If
            Return String.Empty
        End Function

        ' Sets a KeyValuePair in a certain section
        Public Function SetKeyValue(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
            Dim s As IniSection = AddSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.AddKey(sKey)
                If k IsNot Nothing Then
                    k.Value = sValue
                    Return True
                End If
            End If
            Return False
        End Function

        ' Renames an existing section returns true on success, false if the section didn't exist or there was another section with the same sNewSection
        Public Function RenameSection(ByVal sSection As String, ByVal sNewSection As String) As Boolean
            '  Note string trims are done in lower calls.
            Dim bRval As Boolean = False
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                bRval = s.SetName(sNewSection)
            End If
            Return bRval
        End Function

        ' Renames an existing key returns true on success, false if the key didn't exist or there was another section with the same sNewKey
        Public Function RenameKey(ByVal sSection As String, ByVal sKey As String, ByVal sNewKey As String) As Boolean
            '  Note string trims are done in lower calls.
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Dim k As IniSection.IniKey = s.GetKey(sKey)
                If k IsNot Nothing Then
                    Return k.SetName(sNewKey)
                End If
            End If
            Return False
        End Function

        ' Remove a key by section name and key name
        Public Function RemoveKey(ByVal sSection As String, ByVal sKey As String) As Boolean
            Dim s As IniSection = GetSection(sSection)
            If s IsNot Nothing Then
                Return s.RemoveKey(sKey)
            End If
            Return False
        End Function

        ' IniSection class 
        Public Class IniSection
            '  IniFile IniFile object instance
            Private m_pIniFile As IniFile
            '  Name of the section
            Private m_sSection As String
            '  List of IniKeys in the section
            Private m_keys As Hashtable

            ' Constuctor so objects are internally managed
            Protected Friend Sub New(ByVal parent As IniFile, ByVal sSection As String)
                m_pIniFile = parent
                m_sSection = sSection
                m_keys = New Hashtable(StringComparer.InvariantCultureIgnoreCase)
            End Sub

            ' Returns all the keys in a section
            Public ReadOnly Property Keys() As System.Collections.ICollection
                Get
                    Return m_keys.Values
                End Get
            End Property

            ' Returns the section name
            Public ReadOnly Property Name() As String
                Get
                    Return m_sSection
                End Get
            End Property

            ' Adds a key to the IniSection object, returns a IniKey object to the new or existing object
            Public Function AddKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                Dim k As IniSection.IniKey = Nothing
                If sKey.Length <> 0 Then
                    If m_keys.ContainsKey(sKey) Then
                        k = DirectCast(m_keys(sKey), IniKey)
                    Else
                        k = New IniSection.IniKey(Me, sKey)
                        m_keys(sKey) = k
                    End If
                End If
                Return k
            End Function

            ' Removes a single key by string
            Public Function RemoveKey(ByVal sKey As String) As Boolean
                Return RemoveKey(GetKey(sKey))
            End Function

            ' Removes a single key by IniKey object
            Public Function RemoveKey(ByVal Key As IniKey) As Boolean
                If Key IsNot Nothing Then
                    Try
                        m_keys.Remove(Key.Name)
                        Return True
                    Catch ex As Exception
                        Trace.WriteLine(ex.Message)
                    End Try
                End If
                Return False
            End Function

            ' Removes all the keys in the section
            Public Function RemoveAllKeys() As Boolean
                m_keys.Clear()
                Return (m_keys.Count = 0)
            End Function

            ' Returns a IniKey object to the key by name, NULL if it was not found
            Public Function GetKey(ByVal sKey As String) As IniKey
                sKey = sKey.Trim()
                If m_keys.ContainsKey(sKey) Then
                    Return DirectCast(m_keys(sKey), IniKey)
                End If
                Return Nothing
            End Function

            ' Sets the section name, returns true on success, fails if the section
            ' name sSection already exists
            Public Function SetName(ByVal sSection As String) As Boolean
                sSection = sSection.Trim()
                If sSection.Length <> 0 Then
                    ' Get existing section if it even exists...
                    Dim s As IniSection = m_pIniFile.GetSection(sSection)
                    If s IsNot Me AndAlso s IsNot Nothing Then
                        Return False
                    End If
                    Try
                        ' Remove the current section
                        m_pIniFile.m_sections.Remove(m_sSection)
                        ' Set the new section name to this object
                        m_pIniFile.m_sections(sSection) = Me
                        ' Set the new section name
                        m_sSection = sSection
                        Return True
                    Catch ex As Exception
                        Trace.WriteLine(ex.Message)
                    End Try
                End If
                Return False
            End Function

            ' Returns the section name
            Public Function GetName() As String
                Return m_sSection
            End Function

            ' IniKey class
            Public Class IniKey
                '  Name of the Key
                Private m_sKey As String
                '  Value associated
                Private m_sValue As String
                '  Pointer to the parent CIniSection
                Private m_section As IniSection

                ' Constuctor so objects are internally managed
                Protected Friend Sub New(ByVal parent As IniSection, ByVal sKey As String)
                    m_section = parent
                    m_sKey = sKey
                End Sub

                ' Returns the name of the Key
                Public ReadOnly Property Name() As String
                    Get
                        Return m_sKey
                    End Get
                End Property

                ' Sets or Gets the value of the key
                Public Property Value() As String
                    Get
                        Return m_sValue
                    End Get
                    Set(ByVal value As String)
                        m_sValue = value
                    End Set
                End Property

                ' Sets the value of the key
                Public Sub SetValue(ByVal sValue As String)
                    m_sValue = sValue
                End Sub
                ' Returns the value of the Key
                Public Function GetValue() As String
                    Return m_sValue
                End Function

                ' Sets the key name
                ' Returns true on success, fails if the section name sKey already exists
                Public Function SetName(ByVal sKey As String) As Boolean
                    sKey = sKey.Trim()
                    If sKey.Length <> 0 Then
                        Dim k As IniKey = m_section.GetKey(sKey)
                        If k IsNot Me AndAlso k IsNot Nothing Then
                            Return False
                        End If
                        Try
                            ' Remove the current key
                            m_section.m_keys.Remove(m_sKey)
                            ' Set the new key name to this object
                            m_section.m_keys(sKey) = Me
                            ' Set the new key name
                            m_sKey = sKey
                            Return True
                        Catch ex As Exception
                            Trace.WriteLine(ex.Message)
                        End Try
                    End If
                    Return False
                End Function

                ' Returns the name of the Key
                Public Function GetName() As String
                    Return m_sKey
                End Function
            End Class
            ' End of IniKey class
        End Class
        ' End of IniSection class

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class
    ' End of IniFile class
End Module
