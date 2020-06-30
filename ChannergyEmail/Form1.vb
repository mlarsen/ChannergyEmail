Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Check to see if the script is installed in a valid directory


        If Strings.Right(stPath, 1) <> "\" Then
            stPath = stPath + "\"
        End If


        'Get the connection string
        stODBCString = GetDSN(stPath)

        If stODBCString = "" Then
            MessageBox.Show("This utility is not installed in a valid Channergy database folder.", "Error", MessageBoxButtons.OK)
            Close()
        Else
            'Create the EmailHtml tables if necessary
            'Get the version infomration
            Me.txtVersion.Text = "Version:" + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Revision.ToString
            Me.txtPath.Text = stPath

            LoadForm()
        End If
    End Sub

    Private Sub cboEmailName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboEmailName.SelectedIndexChanged
        stEmailName = Me.cboEmailName.SelectedItem
        LoadEmail(stEmailName)
    End Sub


    Private Sub TxtPrompt1_TextChanged(sender As Object, e As EventArgs) Handles TxtPrompt1.TextChanged
        stValue1 = Me.TxtPrompt1.Text
    End Sub

    Private Sub TxtPrompt2_TextChanged(sender As Object, e As EventArgs) Handles TxtPrompt2.TextChanged
        stValue2 = Me.TxtPrompt2.Text
    End Sub

    Private Sub Date1_ValueChanged(sender As Object, e As EventArgs) Handles Date1.ValueChanged
        stValue1 = getDBIASMDate(Date1.Value)
    End Sub

    Private Sub Date2_ValueChanged(sender As Object, e As EventArgs) Handles Date2.ValueChanged
        stValue2 = getDBIASMDate(Date2.Value)
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim stReplace1 As String
        Dim stReplace2 As String


        If bolIsPrompt = True Then
            stCommandSQL = stEmailSQL

            If stValue1 <> "" Then
                If stDataType1 = "DATE" Then
                    stReplace1 = "CAST(" + stValue1 + " AS DATE)"
                ElseIf stDataType1 = "INTEGER" Then
                    stReplace1 = "CAST(" + stValue1 + " AS INTEGER)"
                Else
                    stReplace1 = "CAST('" + stValue1 + "' AS VARCHAR(50))"
                End If
                stCommandSQL = Replace(stEmailSQL, "[PROMPT1]", stReplace1)
            End If

            If stValue2 <> "" Then
                If stDataType2 = "DATE" Then
                    stReplace2 = "CAST(" + stValue2 + " AS DATE)"
                ElseIf stDataType2 = "INTEGER" Then
                    stReplace2 = "CAST(" + stValue2 + " AS INTEGER)"
                Else
                    stReplace2 = "CAST('" + stValue2 + "' AS VARCHAR(50))"
                End If
                stCommandSQL = Replace(stCommandSQL, "[PROMPT2]", stReplace2)
            End If
        Else
            stCommandSQL = stEmailSQL
        End If

        LoadEmailHtmlBatch(stCommandSQL)
        If bolIsError = False Then
            SendEmails()
        End If


    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class
