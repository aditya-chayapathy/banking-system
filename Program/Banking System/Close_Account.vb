Public Class Close_Account
    Private Sub Close_Account_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Button2.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim flag As New Integer
        flag = 0
        If TextBox1.Text.Length = 0 Or TextBox1.Text.Length <> 10 Then
            flag = 1
        ElseIf IsNumeric(TextBox1.Text) = False Then
            flag = 1
        End If
        If flag = 1 Then
            MessageBox.Show("Please Enter Valid Account No !!!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            Dim da As New OleDb.OleDbDataAdapter
            Dim dt As New DataTable
            Dim res As New DialogResult
            Dim y As String
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("SELECT * FROM CUSTOMER WHERE CUS_ACC_NO=?", conn)
            cmd.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
            da.SelectCommand = cmd
            da.Fill(dt)
            DataGridView1.DataSource = dt
            y = TextBox1.Text
            If DataGridView1.RowCount = 2 Then
                Button2.Show()
            Else
                MessageBox.Show("Please Enter Valid Account No !!!", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox1.Text = ""
            End If
            conn.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim res As DialogResult
        res = MessageBox.Show("Are You Sure You Want To Delete This Account  ?", "Delete ?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Global_Variables.random.acc_no = TextBox1.Text
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("DELETE FROM CUSTOMER WHERE CUS_ACC_NO=?", conn)
            cmd.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
            cmd.ExecuteNonQuery()
            Loading_Screen.ShowDialog()
            MessageBox.Show("Account Successfully Closed !!!", "Account Closed", MessageBoxButtons.OK, MessageBoxIcon.None)
            Dim conn2 As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn2.Open()
            Dim cmd6 As New OleDb.OleDbCommand("INSERT INTO TRANSACTION VALUES(?,?,?,?,?,?)", conn2)
            cmd6.Parameters.AddWithValue("?", Global_Variables.random.acc_no)
            cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_name)
            cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_id)
            cmd6.Parameters.AddWithValue("?", String.Format("Account Terminated"))
            cmd6.Parameters.AddWithValue("?", 0)
            cmd6.Parameters.AddWithValue("?", String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now))
            cmd6.ExecuteNonQuery()
            conn2.Close()
            TextBox1.Text = ""
            DataGridView1.DataSource = ""
        Else
            TextBox1.Text = ""
            DataGridView1.DataSource = ""
        End If
    End Sub

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            TextBox1.Text = ""
            DataGridView1.DataSource = ""
            Button2.Hide()
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

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Me.Hide()
        TextBox1.Text = ""
        DataGridView1.DataSource = ""
        Button2.Hide()
        Employee_Menu.Show()
    End Sub

    Private Sub HowToDeleteAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToDeleteAccountToolStripMenuItem.Click
        MessageBox.Show("STEP 1  -  Enter the account no of the customer whose account is to be deleted and then click on the search button" + System.Environment.NewLine + "STEP 2  -  Click the delete button to close the account of the customer", "How To Delete Account ?", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class