Public Class Create_Account

    Private Sub ReturnToPreviousMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToPreviousMenuToolStripMenuItem.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox6.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        DateTimePicker1.ResetText()
        ComboBox1.ResetText()
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        ComboBox5.ResetText()
        ComboBox6.ResetText()
        ComboBox7.ResetText()
        Me.Hide()
        Employee_Menu.Show()
    End Sub

    Private Sub ToolStripLabel3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel3.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox6.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            DateTimePicker1.ResetText()
            ComboBox1.ResetText()
            ComboBox2.ResetText()
            ComboBox3.ResetText()
            ComboBox4.ResetText()
            ComboBox5.ResetText()
            ComboBox6.ResetText()
            ComboBox7.ResetText()
            Me.Hide()
            Log_In_Menu.Show()
            SAPI = CreateObject("SAPI.spvoice")
            SAPI.Speak("Successfully Logged Out")
            Global_Variables.random.log_out_datetime = String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now)
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn.Open()
            Dim cmd2 As New OleDb.OleDbCommand("INSERT INTO LOG_IN VALUES(?,?,?,?)", conn)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.user_name)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.user_id)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.log_in_datetime)
            cmd2.Parameters.AddWithValue("?", Global_Variables.random.log_out_datetime)
            cmd2.ExecuteNonQuery()
            conn.Close()
        Else
            Me.Show()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim flag As New Integer
        flag = 0
        Label18.Text = ""
        Label19.Text = ""
        Label21.Text = ""
        Label20.Text = ""
        Label22.Text = ""
        Label23.Text = ""
        Label24.Text = ""
        Label25.Text = ""
        Label26.Text = ""
        Label27.Text = ""
        Label28.Text = ""
        Label29.Text = ""
        Label30.Text = ""
        Label31.Text = ""
        Label32.Text = ""
        Label33.Text = ""
        Label34.Text = ""
        Label35.Text = ""
        Label36.Text = ""
        Label37.Text = ""
        If TextBox1.Text.Length = 0 Then
            flag = 1
            Label18.Text = "<--"
        End If
        If TextBox2.Text.Length = 0 Then
            flag = 1
            Label19.Text = "<--"
        End If
        If TextBox3.Text.Length = 0 Then
            flag = 1
            Label21.Text = "<--"
        End If
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            flag = 1
            Label23.Text = "<--"
        End If
        If RadioButton3.Checked = False And RadioButton4.Checked = False Then
            flag = 1
            Label26.Text = "<--"
        End If
        If TextBox1.Text.Length > 50 Then
            Label18.Text = "<--"
            flag = 1
        End If
        If TextBox2.Text.Length > 50 Then
            Label19.Text = "<--"
            flag = 1
        End If
        If TextBox3.Text.Length > 50 Then
            Label20.Text = "<--"
            flag = 1
        End If
        If TextBox6.Text.Length > 150 Or TextBox6.Text.Length = 0 Then
            Label28.Text = "<--"
            flag = 1
        End If
        If TextBox10.Text.Length = 0 Or TextBox10.Text.Length <> 10 Then
            Label31.Text = "<--"
            flag = 1
        ElseIf IsNumeric((TextBox10.Text)) = False Then
            Label31.Text = "<--"
            flag = 1
        End If
        If TextBox11.Text.Length = 0 Or TextBox11.Text.Length <> 10 Then
            Label32.Text = "<--"
            flag = 1
        ElseIf IsNumeric((TextBox11.Text)) = False Then
            Label32.Text = "<--"
            flag = 1
        End If
        If TextBox12.Text.Length = 0 Or Not (TextBox12.Text.Contains("@")) Then
            Label33.Text = "<--"
            flag = 1
        End If
        If TextBox13.Text.Length = 0 Then
            Label34.Text = "<--"
            flag = 1
        ElseIf IsNumeric((TextBox13.Text)) = False Then
            Label34.Text = "<--"
            flag = 1
        End If
        If DateTimePicker1.Value.Date = DateTime.Now.Date Then
            flag = 1
            Label22.Text = "<--"
        End If
        If ComboBox1.Text.Length = 0 Then
            flag = 1
            Label24.Text = "<--"
        End If
        If ComboBox2.Text.Length = 0 Then
            flag = 1
            Label25.Text = "<--"
        End If
        If ComboBox3.Text.Length = 0 Then
            flag = 1
            Label27.Text = "<--"
        End If
        If ComboBox4.Text.Length = 0 Then
            flag = 1
            Label35.Text = "<--"
        End If
        If ComboBox5.Text.Length = 0 Then
            flag = 1
            Label36.Text = "<--"
        End If
        If ComboBox6.Text.Length = 0 Then
            flag = 1
            Label29.Text = "<--"
        End If
        If ComboBox7.Text.Length = 0 Then
            flag = 1
            Label30.Text = "<--"
        End If
        If flag = 1 Then
            Label37.Text = "'<--'     INDICATES CONSTRAINT VOILATION OF THE FIELD"
        Else
            Dim gender As String
            Dim senior_citizen As String
            Dim acc_no As New Double
            Dim acc_update As New Double
            If RadioButton1.Checked = True Then
                gender = "Male"
            Else
                gender = "Female"
            End If
            If RadioButton3.Checked = True Then
                senior_citizen = "Yes"
            Else
                senior_citizen = "No"
            End If
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            Dim da As New OleDb.OleDbDataAdapter
            Dim dt As New DataTable
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM ACC_NO", conn)
            da.SelectCommand = cmd
            da.Fill(dt)
            acc_no = dt.Rows(0).Item(0)
            acc_update = acc_no + 1
            Dim cmd1 As New OleDb.OleDbCommand("UPDATE ACC_NO SET NO=? WHERE NO=?", conn)
            cmd1.Parameters.AddWithValue("?", acc_update)
            cmd1.Parameters.AddWithValue("?", acc_no)
            cmd1.ExecuteNonQuery()
            Global_Variables.random.acc_no = acc_no
            Dim cmd2 As New OleDb.OleDbCommand("INSERT INTO CUSTOMER VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", conn)
            cmd2.Parameters.AddWithValue("?", TextBox1.Text)
            cmd2.Parameters.AddWithValue("?", TextBox2.Text)
            cmd2.Parameters.AddWithValue("?", TextBox3.Text)
            cmd2.Parameters.AddWithValue("?", DateTimePicker1.Text)
            cmd2.Parameters.AddWithValue("?", gender)
            cmd2.Parameters.AddWithValue("?", ComboBox3.Text)
            cmd2.Parameters.AddWithValue("?", senior_citizen)
            cmd2.Parameters.AddWithValue("?", TextBox6.Text)
            cmd2.Parameters.AddWithValue("?", CDbl(TextBox10.Text))
            cmd2.Parameters.AddWithValue("?", TextBox12.Text)
            cmd2.Parameters.AddWithValue("?", CDbl(TextBox11.Text))
            cmd2.Parameters.AddWithValue("?", ComboBox1.Text)
            cmd2.Parameters.AddWithValue("?", ComboBox2.Text)
            cmd2.Parameters.AddWithValue("?", CDbl(TextBox13.Text))
            cmd2.Parameters.AddWithValue("?", ComboBox4.Text)
            cmd2.Parameters.AddWithValue("?", ComboBox5.Text)
            cmd2.Parameters.AddWithValue("?", acc_no)
            cmd2.Parameters.AddWithValue("?", 0)
            cmd2.Parameters.AddWithValue("?", Date.Now())
            cmd2.Parameters.AddWithValue("?", ComboBox7.Text)
            cmd2.Parameters.AddWithValue("?", ComboBox6.Text)
            cmd2.ExecuteNonQuery()
            conn.Close()
            Loading_Screen.ShowDialog()
            Dim bailout As New DateTime
            bailout = DateTime.Now.AddMilliseconds(2500)
            While (DateTime.Now < bailout)
            End While
            MessageBox.Show("Customer Successfully Added" + System.Environment.NewLine + "Customer Account No  =  " + acc_no.ToString(), "Account Created !!!", MessageBoxButtons.OK, MessageBoxIcon.None)
            Dim conn2 As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn2.Open()
            Dim cmd6 As New OleDb.OleDbCommand("INSERT INTO TRANSACTION VALUES(?,?,?,?,?,?)", conn2)
            cmd6.Parameters.AddWithValue("?", Global_Variables.random.acc_no)
            cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_name)
            cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_id)
            cmd6.Parameters.AddWithValue("?", String.Format("New Account"))
            cmd6.Parameters.AddWithValue("?", 0)
            cmd6.Parameters.AddWithValue("?", String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now))
            cmd6.ExecuteNonQuery()
            conn2.Close()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox6.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            DateTimePicker1.ResetText()
            ComboBox1.ResetText()
            ComboBox2.ResetText()
            ComboBox3.ResetText()
            ComboBox4.ResetText()
            ComboBox5.ResetText()
            ComboBox6.ResetText()
            ComboBox7.ResetText()
            Me.Hide()
            Employee_Menu.Show()
        End If
    End Sub

    Private Sub Create_Account_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        Dim dr As OleDb.OleDbDataReader
        conn.Open()
        Dim cmd1 As New OleDb.OleDbCommand("SELECT ACC_NAME FROM ACCOUNT_TYPE", conn)
        dr = cmd1.ExecuteReader
        While dr.Read()
            ComboBox7.Items.Add(dr.GetValue(0))
        End While
        Dim cmd2 As New OleDb.OleDbCommand("SELECT BR_NAME FROM BRANCH", conn)
        dr = cmd2.ExecuteReader
        While dr.Read()
            ComboBox6.Items.Add(dr.GetValue(0))
        End While
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "dd MMM yyyy"
        conn.Close()
    End Sub

    Private Sub HowToCreateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToCreateToolStripMenuItem.Click
        MessageBox.Show("Fill in all details provided by the customer." + System.Environment.NewLine + "Click on 'Submit' button to create the account.", "Create Account ?", MessageBoxButtons.OK)
    End Sub

    Private Sub DataEntryConstraintsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataEntryConstraintsToolStripMenuItem.Click
        MessageBox.Show("Name - Shouldnt Exceed 50 Characters" + System.Environment.NewLine + "Father's Name - Shouldnt Exceed 50 characters" + System.Environment.NewLine + "Mother's Name - Shouldnt Exceed 50 Characters" + System.Environment.NewLine + "Address - Shouldnt Exceed More Than 150 Characters" + System.Environment.NewLine + "Contact No - Entered Text Must Contain Only 10 Numbers" + System.Environment.NewLine + "Aadhar ID - Entered Text Must Contain Only 10 Numbers", "Constraints On Fields  ?", MessageBoxButtons.OK)
    End Sub
End Class