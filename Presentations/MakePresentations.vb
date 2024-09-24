Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common

Public Class MakePresentations
    Private ds, ds_server As New DataSet

    Private slidenr As Integer = 1
    Private subjectnr As Integer = 1



    Private Sub frmAccounts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Dim f As New Font("arial", 12, FontStyle.Bold)
        'Me.BackColor = Color.FromName("ActiveCaption")


    End Sub

    Private Sub FormatColumnHeaders()

        Dim ts As New DataGridTableStyle
        Dim cs1, cs2, cs3, cs4, cs5, cs6, cs7, cs8, cs9 As DataGridTextBoxColumn

        ts.MappingName = "Presentation"

        'select presnr,slide,subject,action,type,size1,num,text


        cs1 = New DataGridTextBoxColumn
        cs1.MappingName = "slide"
        cs1.HeaderText = "Slide:"
        cs1.Width = 50
        ts.GridColumnStyles.Add(cs1)

        cs2 = New DataGridTextBoxColumn
        cs2.MappingName = "subject"
        cs2.HeaderText = "Subject:"
        cs2.Width = 50
        ts.GridColumnStyles.Add(cs2)

        cs3 = New DataGridTextBoxColumn
        cs3.MappingName = "action"
        cs3.HeaderText = "Action:"
        cs3.Width = 50
        ts.GridColumnStyles.Add(cs3)

        cs4 = New DataGridTextBoxColumn
        cs4.MappingName = "type"
        cs4.HeaderText = "Type:"
        cs4.Width = 50
        ts.GridColumnStyles.Add(cs4)

        cs5 = New DataGridTextBoxColumn
        cs5.MappingName = "size1"
        cs5.HeaderText = "Align:"
        cs5.Width = 50
        ts.GridColumnStyles.Add(cs5)

        cs6 = New DataGridTextBoxColumn
        cs6.MappingName = "num"
        cs6.HeaderText = "Amount:"
        cs6.Width = 50
        ts.GridColumnStyles.Add(cs6)

        cs8 = New DataGridTextBoxColumn
        cs8.MappingName = "text"
        cs8.HeaderText = "Text:"
        cs8.Width = 380
        ts.GridColumnStyles.Add(cs8)






        DataGrid1.TableStyles.Add(ts)

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim myrow As DataRow = Me.ds.Tables(0).NewRow()
        ds.Tables(0).Rows.Add(myrow)

        DataGrid1.CurrentRowIndex = ds.Tables(0).Rows.Count()
        DataGrid1.Select(ds.Tables(0).Rows.Count() - 1)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim myrow As DataRow = Me.ds.Tables(0).Rows(DataGrid1.CurrentRowIndex)
        myrow.Delete()

    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.MouseEnter


    End Sub

    Private Sub Button1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.MouseLeave


    End Sub

    Private Sub Button2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseEnter


    End Sub

    Private Sub Button2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseLeave

    End Sub

    Private Sub Button3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.MouseEnter



    End Sub

    Private Sub Button3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.MouseLeave



    End Sub

    Private Sub Button4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.MouseEnter


    End Sub

    Private Sub Button4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.MouseLeave


    End Sub



    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click


        PresentationNr = CInt(TextBox10.Text)
        Me.LoadPresentations()
        Label1.Visible = False
        TextBox10.Visible = False
        Button6.Visible = False

        DataGrid1.Height = DataGrid1.Height + 30
        DataGrid1.Location = New Point(DataGrid1.Location.X, DataGrid1.Location.Y - 30)


    End Sub

    Private Sub LoadPresentations()

        Dim srv, user, pwd, db, mysql, type As String
        Dim mystring As String
        Dim myconnection1 As DbConnection
        Dim da As DbDataAdapter

        ds_server.ReadXml(CurDir() + "\xml\server.xml")
        ds.Reset()

        Me.FormatColumnHeaders()

        ' Open connection
        type = ds_server.Tables(0).Rows(0).Item(0)
        srv = ds_server.Tables(0).Rows(0).Item(1)
        user = ds_server.Tables(0).Rows(0).Item(2)
        pwd = ds_server.Tables(0).Rows(0).Item(3)
        db = ds_server.Tables(0).Rows(0).Item(4)


        mystring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + db + "';Jet OLEDB:Database Password=Tapyxe01"
        myconnection1 = New OleDbConnection(mystring)
        'mysql = "select nr,name,amount,length from Fishes where bread = 'x' order by nr"
        mysql = "select presnr,slide,subject,action,type,size1,num,text from Presentation where presnr = " + PresentationNr.ToString() + " order by presnr,slide,subject"

        da = New OleDbDataAdapter
        da.SelectCommand = New OleDbCommand(mysql, myconnection1)

        Try
            myconnection1.Open()
            da.Fill(ds, "Presentation")
            DataGrid1.DataSource = ds.Tables(0)

        Catch ex As Exception
            MsgBox(ex.Message)


        Finally
            myconnection1.Close()

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim my_sql As String
        Dim srv, user, pwd, db, type As String
        Dim mystring As String
        Dim i As Integer
        Dim myconnection1 As DbConnection
        Dim mycommand As DbCommand

        ' Open connection
        type = ds_server.Tables(0).Rows(0).Item(0)
        srv = ds_server.Tables(0).Rows(0).Item(1)
        user = ds_server.Tables(0).Rows(0).Item(2)
        pwd = ds_server.Tables(0).Rows(0).Item(3)
        db = ds_server.Tables(0).Rows(0).Item(4)

        Select Case type
            Case "ACCESS"
                mystring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + db + "';Jet OLEDB:Database Password=Tapyxe01"
                myconnection1 = New OleDbConnection(mystring)
                mycommand = New OleDbCommand

        End Select

        Try
            myconnection1.Open()
            mycommand.Connection = myconnection1
            my_sql = "delete from Presentation where presnr = " + PresentationNr.ToString()
            mycommand.CommandText = my_sql
            mycommand.ExecuteNonQuery()

            ds.AcceptChanges()

            'presnr,slide,subject,action,type,size1,num,text

            For i = 0 To ds.Tables(0).Rows.Count - 1
                my_sql = "insert into Presentation values ("
                my_sql = my_sql + PresentationNr.ToString() + ","
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(1).ToString() + ","
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(2).ToString() + ",'"
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(3) + "','"
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(4) + "','"
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(5) + "',"
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(6).ToString() + ",'"
                my_sql = my_sql + ds.Tables(0).Rows(i).Item(7) + "')"
                mycommand.CommandText = my_sql
                mycommand.ExecuteNonQuery()
            Next

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally
            myconnection1.Close()

        End Try

    End Sub


    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click

        CheckBox1.Checked = True
        CheckBox2.Checked = False


    End Sub


    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click

        CheckBox1.Checked = False
        CheckBox2.Checked = True

    End Sub


    Private Sub CheckBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.Click

        CheckBox4.Checked = False
        CheckBox3.Checked = True

    End Sub


    Private Sub CheckBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox4.Click

        CheckBox4.Checked = True
        CheckBox3.Checked = False

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        slidenr = slidenr + 1
        TextBox1.Text = slidenr.ToString()
        subjectnr = 1
        TextBox2.Text = subjectnr.ToString()


    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click


        Dim myrow As DataRow = Me.ds.Tables(0).NewRow()
        myrow.Item(0) = PresentationNr
        myrow.Item(1) = CInt(TextBox1.Text)
        myrow.Item(2) = CInt(TextBox2.Text)
        If CheckBox1.Checked = True Then
            myrow.Item(3) = "D"
        Else
            myrow.Item(3) = "S"
        End If
        myrow.Item(4) = TextBox3.Text
        If CheckBox4.Checked = True Then
            myrow.Item(5) = "C"
        Else
            myrow.Item(5) = "L"
        End If
        myrow.Item(6) = CInt(TextBox4.Text)
        ds.Tables(0).Rows.Add(myrow)

        DataGrid1.CurrentRowIndex = ds.Tables(0).Rows.Count()
        DataGrid1.Select(ds.Tables(0).Rows.Count() - 1)

        subjectnr = subjectnr + 1
        TextBox2.Text = subjectnr.ToString()



    End Sub
End Class