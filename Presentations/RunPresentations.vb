
Imports System
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Common

Public Class RunPresentations

    Private NumItems As Integer = 0
    Private TitleHeight As Integer = 30
    Private EndSlides As Boolean = False

    Private FirstHeight As Integer = 150
    Private FirstHeightSmall As Integer = 100

    Private SmallInterval As Integer = 70
    Private FirstInterval As Integer = 100
    Private LargeInterval As Integer = 200

    Private posxStop As Integer = 1200
    Private posyStop As Integer = 1000


    Private LineNum As Integer = 1
    Private Cellrow As Integer = 2
    Private SlideNumber As Integer = 1

    Private MySlide(1000) As Integer
    Private MySubject(1000) As Integer
    Private MyAction(1000) As String
    Private MyType(1000) As String
    Private MySize(1000) As String
    Private MyNumbers(1000) As String

    Private MyTexts(1000) As String


    Private myfile, mydir, CurDirectory, InitialDir, ExcelFileName As String
    Private CurCell, CurCell2, CurCell3, CurCell4, CurCell5, CurCell6, CurCell7 As String

    Private ds, ds_server As New DataSet

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean

        If keyData = Keys.Enter Or keyData = Keys.Space Or keyData = Keys.Return Then 'Fire Bullet

            If LineNum > ds.Tables(0).Rows.Count Then

                Timer1.Stop()
                MsgBox("This is the end of this presentation!")
                Me.Close()
            Else
                Timer1.Start()
            End If



            Return True
        End If

        If keyData = Keys.S Or Keys.Escape Then
            EndSlides = True
        End If

    End Function




    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        OpenFileDialog1.InitialDirectory = mydir
        OpenFileDialog1.Filter = "Excel Files|*.xls"

        CurDirectory = CurDir()
        InitialDir = CurDirectory
        'For i = 0 To 1000
        '    MySlide(i) = 0
        '    MyNumbers(i) = 0
        'Next

    End Sub

    Private Sub LoadPresentations()

        Dim srv, user, pwd, db, mysql, type As String
        Dim mystring As String
        Dim myconnection1 As DbConnection
        Dim da As DbDataAdapter

        ds_server.ReadXml(CurDir() + "\xml\server.xml")
        ds.Reset()

        'Me.FormatColumnHeaders()

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
            LogError(ex.Message, "Presentation LOAD")

        Finally
            myconnection1.Close()

        End Try


    End Sub



    Public Sub LogError(ByVal msg As String, ByVal filnam As String)

        Dim myfile As String
        myfile = CurDir() + "\log\error.logging"

        Dim FS As New FileStream(myfile, FileMode.Open, FileAccess.ReadWrite)
        Dim SR As New StreamWriter(FS)

        FS.Seek(0, SeekOrigin.End)
        SR.WriteLine("Execution Date: " + Now() + " , Program: " + filnam)
        SR.WriteLine("Error: " + msg)

        SR.Close()
        FS.Close()


    End Sub


    Private Sub PresentText(ByVal Subjekt As Integer, ByVal Action As String, ByVal Type As String, ByVal Size As String, ByVal MyText As String, ByVal h1 As Integer, ByVal delta As Integer)

        Dim MyTextBox As TextBox
        Dim MyHeight As Integer
        Dim str, str2, str3 As String


        Select Case Subjekt
            Case 1
                MyTextBox = TextBox0
                MyHeight = TitleHeight
            Case 2
                MyTextBox = TextBox1
                MyHeight = h1
            Case 3
                MyTextBox = TextBox2
                MyHeight = h1 + delta
            Case 4
                MyTextBox = TextBox3
                MyHeight = h1 + (delta * 2)
            Case 5
                MyTextBox = TextBox4
                MyHeight = h1 + (delta * 3)
            Case 6
                MyTextBox = TextBox5
                MyHeight = h1 + (delta * 4)
            Case 7
                MyTextBox = TextBox6
                MyHeight = h1 + (delta * 5)
            Case 8
                MyTextBox = TextBox7
                MyHeight = h1 + (delta * 6)
            Case 9
                MyTextBox = TextBox8
                MyHeight = h1 + (delta * 7)

        End Select

        MyTextBox.Text = MyText
        str = ComboBox1.Text
        If Subjekt = 1 Then
            MyTextBox.Font = New Font(str, 24, FontStyle.Bold)
        Else
            MyTextBox.Font = New Font(str, 20, FontStyle.Bold)
        End If

        'backcolour
        str3 = ComboBox3.Text
        Me.BackColor = Color.FromName(str3)
        MyTextBox.BackColor = Me.BackColor

        'forecolour
        str2 = ComboBox2.Text
        MyTextBox.ForeColor = Color.FromName(str2)

        MyTextBox.Top = MyHeight
        MyTextBox.Visible = True
        If Size = "C" Then
            MyTextBox.TextAlign = HorizontalAlignment.Center
        Else
            MyTextBox.TextAlign = HorizontalAlignment.Left
        End If

    End Sub

    Private Sub ClearTextBox()

        TextBox0.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""

        TextBox0.Visible = False
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        TextBox6.Visible = False
        TextBox7.Visible = False
        TextBox8.Visible = False

    End Sub

    Private Sub PresentSlides()

        Dim CurSlide As Integer
        Dim CurSubject As Integer
        Dim CurAction As String
        Dim CurType As String
        Dim CurSize As String
        Dim CurText As String
        Dim CurNum, y1, delta As Integer

        Dim LineNum As Integer = 1
        Dim TextNum As Integer = 1
        Dim TheInput As String

        LineNum = CInt(TextBox10.Text)
        CurSlide = MySlide(LineNum)
        'While CurSlide = SlideNumber
        While CurSlide > 0 And Not EndSlides = True
            CurSubject = MySubject(LineNum)
            CurAction = MyAction(LineNum)
            CurType = MyType(LineNum)
            CurSize = MySize(LineNum)
            CurText = MyTexts(LineNum)
            If MyNumbers(LineNum) > 0 Then
                CurNum = MyNumbers(LineNum)
            End If
            If CurNum <= 5 Then
                y1 = FirstHeight
                delta = FirstInterval
            Else
                y1 = FirstHeightSmall
                delta = SmallInterval
            End If
            Me.PresentText(CurSubject, CurAction, CurType, CurSize, CurText, y1, delta)
            If CurAction = "S" Then
                Me.Refresh()
                TheInput = Microsoft.VisualBasic.InputBox("Next:", "", "", posxStop, posyStop)
                If TheInput.ToUpper() = "STOP" Then
                    'Me.Close()
                    EndSlides = True
                End If
            End If
            TextNum = TextNum + 1
            LineNum = LineNum + 1
            CurSlide = MySlide(LineNum)
            If CurSlide > SlideNumber Then 'nieuwe slide
                Me.ClearTextBox()
                SlideNumber = SlideNumber + 1
            End If

        End While

    End Sub


    Private Sub AdjustCells()
        Dim cell, newcell As String

        'Adjust Cell Rows
        Cellrow = Cellrow + 1
        'Get 
        CurCell = "A" + Cellrow.ToString()
        CurCell2 = "B" + Cellrow.ToString()
        CurCell3 = "C" + Cellrow.ToString()
        CurCell4 = "D" + Cellrow.ToString()
        CurCell5 = "E" + Cellrow.ToString()
        CurCell6 = "F" + Cellrow.ToString()
        CurCell7 = "G" + Cellrow.ToString()

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        EndSlides = True
        Me.Close()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Me.PresentSlidesDB()

    End Sub

    Private Sub PresentSlidesDB()

        Dim CurSlide As Integer
        Dim CurSubject As Integer
        Dim CurAction As String
        Dim CurType As String
        Dim CurSize As String
        Dim CurText As String
        Dim CurNum, y1, delta As Integer

        Dim LineNum As Integer = 1
        Dim TextNum As Integer = 1
        Dim TheInput As String


        LineNum = CInt(TextBox10.Text)
        CurSlide = ds.Tables(0).Rows(LineNum - 1).Item(1)

        'CurCell = "A2"
        'CurCell2 = "B2"
        'CurCell3 = "C2"
        'CurCell4 = "D2"
        'CurCell5 = "E2"
        'CurCell6 = "F2"
        'CurCell7 = "G2"
        'CurSlide = xlRange.Value
        'CurSubject = xlRange2.Value
        'CurAction = xlRange3.Value
        'CurType = xlRange4.Value
        'CurSize1 = xlRange5.Value
        'CurNum = xlRange6.Value
        'CurText = xlRange7.Value

        'While CurSlide = SlideNumber
        While CurSlide > 0 And Not EndSlides = True
            CurSubject = ds.Tables(0).Rows(LineNum - 1).Item(2)
            CurAction = ds.Tables(0).Rows(LineNum - 1).Item(3)
            CurType = ds.Tables(0).Rows(LineNum - 1).Item(4)
            CurSize = ds.Tables(0).Rows(LineNum - 1).Item(5)
            CurNum = ds.Tables(0).Rows(LineNum - 1).Item(6)
            CurText = ds.Tables(0).Rows(LineNum - 1).Item(7)
            If CurNum <= 5 Then
                y1 = FirstHeight
                delta = FirstInterval
            Else
                y1 = FirstHeightSmall
                delta = SmallInterval
            End If

            Me.PresentText(CurSubject, CurAction, CurType, CurSize, CurText, y1, delta)
            If CurAction = "S" Then
                Me.Refresh()
                TheInput = Microsoft.VisualBasic.InputBox("Next:", "", "", posxStop, posyStop)
                If TheInput.ToUpper() = "STOP" Then
                    'Me.Close()
                    EndSlides = True
                End If
            End If
            TextNum = TextNum + 1
            LineNum = LineNum + 1
            'CurSlide = MySlide(LineNum)
            CurSlide = ds.Tables(0).Rows(LineNum - 1).Item(1)
            If CurSlide > SlideNumber Then 'nieuwe slide
                Me.ClearTextBox()
                SlideNumber = SlideNumber + 1
            End If

        End While

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim CurSlide As Integer
        Dim CurSubject As Integer
        Dim CurAction As String
        Dim CurType As String
        Dim CurSize As String
        Dim CurText As String
        Dim CurNum, y1, delta As Integer

        Dim TextNum As Integer = 1
        Dim TheInput As String

        CurSlide = ds.Tables(0).Rows(LineNum - 1).Item(1)
        If CurSlide > SlideNumber Then 'nieuwe slide
            Me.ClearTextBox()
            SlideNumber = SlideNumber + 1
        End If

        CurSubject = ds.Tables(0).Rows(LineNum - 1).Item(2)
        CurAction = ds.Tables(0).Rows(LineNum - 1).Item(3)
        CurType = ds.Tables(0).Rows(LineNum - 1).Item(4)
        CurSize = ds.Tables(0).Rows(LineNum - 1).Item(5)
        CurNum = ds.Tables(0).Rows(LineNum - 1).Item(6)
        CurText = ds.Tables(0).Rows(LineNum - 1).Item(7)
        If CurNum <= 5 Then
            y1 = FirstHeight
            delta = FirstInterval
        Else
            y1 = FirstHeightSmall
            delta = SmallInterval
        End If

        Me.PresentText(CurSubject, CurAction, CurType, CurSize, CurText, y1, delta)
        If CurAction = "S" Then
            Me.Refresh()
            Timer1.Stop()
        End If
        TextNum = TextNum + 1
        LineNum = LineNum + 1




    End Sub

    Private Sub HideControls()

        Button4.Visible = False
        Label1.Visible = False
        TextBox10.Visible = False

        ComboBox1.Visible = False
        ComboBox2.Visible = False
        ComboBox3.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        TextBox11.Visible = False


    End Sub

    Private Sub ApplyText()
        Dim str As String

        str = ComboBox1.Text
        If str <> "Microsoft Sans Serif" Then


        End If

        TextBox11.Font = New Font(str, 11, FontStyle.Bold)

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim str As String

        LineNum = 1
        PresentationNr = CInt(TextBox10.Text)
        Me.LoadPresentations()
        Me.HideControls()

        Timer1.Start()


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim str As String

        str = ComboBox1.Text
        TextBox11.Font = New Font(str, 11, FontStyle.Bold)


    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

        Dim str As String

        str = ComboBox3.Text
        TextBox11.BackColor = Color.FromName(str)

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged


        Dim str As String

        str = ComboBox2.Text
        TextBox11.ForeColor = Color.FromName(str)


    End Sub
End Class