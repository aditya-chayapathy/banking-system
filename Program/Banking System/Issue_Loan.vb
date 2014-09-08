Public Class Issue_Loan

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            TextBox1.Text = ""
            TextBox2.Text = ""
            ComboBox1.ResetText()
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

    Private Sub ReturnToPreviousMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToPreviousMenuToolStripMenuItem.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.ResetText()
        Me.Hide()
        Employee_Menu.Show()
    End Sub

    Private Sub HowToIssueLoanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToIssueLoanToolStripMenuItem.Click
        MessageBox.Show("First, fill all the specified fields. Next, Click on 'Submit' button to issue loan" + System.Environment.NewLine + "NOTE  -  A CUSTOMER CANNOT BE GRANTED MULTIPLE LOANS AT A TIME", "How To Issue Loan ?", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Issue_Loan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DataGridView1.Hide()
        Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
        Dim dr As OleDb.OleDbDataReader
        conn.Open()
        Dim cmd1 As New OleDb.OleDbCommand("SELECT L_NAME FROM LOAN_TYPE", conn)
        dr = cmd1.ExecuteReader
        While dr.Read()
            ComboBox1.Items.Add(dr.GetValue(0))
        End While
        conn.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim flag As New Integer
        flag = 0
        If TextBox1.Text.Length = 0 Or TextBox1.Text.Length <> 10 Then
            flag = 1
        ElseIf IsNumeric(TextBox1.Text) = False Then
            flag = 1
        End If
        If TextBox2.Text.Length = 0 Then
            flag = 1
        ElseIf IsNumeric(TextBox2.Text) = False Then
            flag = 1
        End If
        If ComboBox1.Text = "" Then
            flag = 1
        End If
        If flag = 1 Then
            MessageBox.Show("Please fill all the fields correctly to proceed.....", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Global_Variables.random.acc_no = TextBox1.Text
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            Dim dr As OleDb.OleDbDataReader
            Dim da As New OleDb.OleDbDataAdapter
            Dim dt As New DataTable
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM CUSTOMER WHERE CUS_ACC_NO=?", conn)
            cmd.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
            da.SelectCommand = cmd
            da.Fill(dt)
            DataGridView1.DataSource = dt
            If DataGridView1.RowCount <> 2 Then
                MessageBox.Show("Please Enter Valid Account No !!!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox1.Text = ""
                TextBox2.Text = ""
                ComboBox1.ResetText()
                conn.Close()
            Else
                Dim cmd1 As New OleDb.OleDbCommand("SELECT * FROM BORROWER WHERE BO_ACC_NO=?", conn)
                cmd1.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                dr = cmd1.ExecuteReader
                If dr.Read() Then
                    Loading_Screen.ShowDialog()
                    MessageBox.Show("This customer has already taken a loan" + System.Environment.NewLine + "MULTIPLE LOANS CANNOT BE SANCTIONED", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    ComboBox1.ResetText()
                    conn.Close()
                Else
                    Dim cmd2 As New OleDb.OleDbCommand("INSERT INTO BORROWER VALUES(?,?,?)", conn)
                    cmd2.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                    cmd2.Parameters.AddWithValue("?", ComboBox1.Text)
                    cmd2.Parameters.AddWithValue("?", CDbl(TextBox2.Text))
                    cmd2.ExecuteNonQuery()
                    Loading_Screen.ShowDialog()
                    MessageBox.Show("Loan Successfully Granted !!!", "Loan Granted", MessageBoxButtons.OK, MessageBoxIcon.None)
                    Dim conn2 As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
                    conn2.Open()
                    Dim cmd6 As New OleDb.OleDbCommand("INSERT INTO TRANSACTION VALUES(?,?,?,?,?,?)", conn2)
                    cmd6.Parameters.AddWithValue("?", Global_Variables.random.acc_no)
                    cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_name)
                    cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_id)
                    cmd6.Parameters.AddWithValue("?", String.Format("Loan Issued"))
                    cmd6.Parameters.AddWithValue("?", TextBox2.Text)
                    cmd6.Parameters.AddWithValue("?", String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now))
                    cmd6.ExecuteNonQuery()
                    conn2.Close()
                    TextBox1.Text = ""
                    TextBox2.Text = ""
                    ComboBox1.ResetText()
                End If
            End If
        End If
    End Sub
End Class