Public Class Transactions

    Private Sub Transactions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.ResetText()
        DataGridView2.DataSource = ""
        DataGridView3.DataSource = ""
        DataGridView1.DataSource = ""
        DataGridView1.Hide()
        DataGridView2.Hide()
        DataGridView3.Hide()
        Label6.Hide()
        Label5.Hide()
        Label4.Hide()
    End Sub

    Private Sub ReturnToPreviousMenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReturnToPreviousMenuToolStripMenuItem.Click
        Me.Hide()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox1.ResetText()
        DataGridView2.DataSource = ""
        DataGridView3.DataSource = ""
        DataGridView1.DataSource = ""
        DataGridView1.Hide()
        DataGridView2.Hide()
        DataGridView3.Hide()
        Label6.Hide()
        Label5.Hide()
        Label4.Hide()
        Employee_Menu.Show()
    End Sub

    Private Sub HowToDepositToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToDepositToolStripMenuItem.Click
        MessageBox.Show("STEP 1  -  Enter account no and amount." + System.Environment.NewLine + "STEP 2  -  Choose 'Deposit' under transaction option" + System.Environment.NewLine + "STEP 3  -  Click on 'Submit' button", "How To Deposit ?", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub HowToWithdrawToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToWithdrawToolStripMenuItem.Click
        MessageBox.Show("STEP 1  -  Enter account no and amount." + System.Environment.NewLine + "STEP 2  -  Choose 'Withdraw' under transaction option" + System.Environment.NewLine + "STEP 3  -  Click on 'Submit' button", "How To Deposit ?", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            TextBox1.Text = ""
            TextBox2.Text = ""
            ComboBox1.ResetText()
            DataGridView2.DataSource = ""
            DataGridView3.DataSource = ""
            DataGridView1.DataSource = ""
            DataGridView1.Hide()
            DataGridView2.Hide()
            DataGridView3.Hide()
            Label6.Hide()
            Label5.Hide()
            Label4.Hide()
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
        If ComboBox1.Text.Length = 0 Then
            flag = 1
        End If
        If flag = 1 Then
            MessageBox.Show("Please Fill All the Fields Correctly To Proceed.......", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Global_Variables.random.acc_no = TextBox1.Text
            Dim conn As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
            conn.Open()
            Dim cmd As New OleDb.OleDbCommand("SELECT CUS_NAME,CUS_ACC_NO,CUS_BAL,CUS_ACC_TYPE,CUS_BRANCH FROM CUSTOMER WHERE CUS_ACC_NO=? ", conn)
            cmd.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
            Dim da As New OleDb.OleDbDataAdapter
            da.SelectCommand = cmd
            Dim dt As New DataTable
            da.Fill(dt)
            DataGridView2.DataSource = dt
            If DataGridView2.RowCount = 2 Then
                Label4.Show()
                DataGridView2.Show()
                Label6.Show()
                Label5.Show()
                If ComboBox1.Text = "Deposit" Then
                    Loading_Screen.ShowDialog()
                    Dim cmd1 As New OleDb.OleDbCommand("UPDATE CUSTOMER SET CUS_BAL=CUS_BAL+? WHERE CUS_ACC_NO=?", conn)
                    cmd1.Parameters.AddWithValue("?", CDbl(TextBox2.Text))
                    cmd1.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                    cmd1.ExecuteNonQuery()
                    Dim cmd7 As New OleDb.OleDbCommand("SELECT CUS_BAL FROM CUSTOMER WHERE CUS_ACC_NO=?", conn)
                    cmd7.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                    da.SelectCommand = cmd7
                    da.Fill(dt)
                    DataGridView1.DataSource = dt
                    Dim conn2 As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
                    conn2.Open()
                    Dim cmd6 As New OleDb.OleDbCommand("INSERT INTO TRANSACTION VALUES(?,?,?,?,?,?)", conn2)
                    cmd6.Parameters.AddWithValue("?", Global_Variables.random.acc_no)
                    cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_name)
                    cmd6.Parameters.AddWithValue("?", Global_Variables.random.user_id)
                    cmd6.Parameters.AddWithValue("?", String.Format("Deposit"))
                    cmd6.Parameters.AddWithValue("?", TextBox2.Text)
                    cmd6.Parameters.AddWithValue("?", String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now))
                    cmd6.ExecuteNonQuery()
                    conn2.Close()
                    conn.Close()
                ElseIf ComboBox1.Text = "Withdraw" Then
                    Dim cmd5 As New OleDb.OleDbCommand("SELECT CUS_BAL FROM CUSTOMER WHERE CUS_ACC_NO=?", conn)
                    cmd5.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                    da.SelectCommand = cmd5
                    da.Fill(dt)
                    DataGridView3.DataSource = dt
                    If DataGridView3(2, 0).Value > CDbl(TextBox2.Text) Then
                        Loading_Screen.ShowDialog()
                        Dim cmd3 As New OleDb.OleDbCommand("UPDATE CUSTOMER SET CUS_BAL=CUS_BAL-? WHERE CUS_ACC_NO=?", conn)
                        cmd3.Parameters.AddWithValue("?", CDbl(TextBox2.Text))
                        cmd3.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                        cmd3.ExecuteNonQuery()
                        Dim cmd6 As New OleDb.OleDbCommand("SELECT CUS_BAL FROM CUSTOMER WHERE CUS_ACC_NO=?", conn)
                        cmd6.Parameters.AddWithValue("?", CDbl(TextBox1.Text))
                        da.SelectCommand = cmd6
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                        conn.Close()
                        Dim conn2 As New OleDb.OleDbConnection("Provider=MSDAORA;Data Source=aditya;User ID=aditya; Password=aditya;")
                        conn2.Open()
                        Dim cmd7 As New OleDb.OleDbCommand("INSERT INTO TRANSACTION VALUES(?,?,?,?,?,?)", conn2)
                        cmd7.Parameters.AddWithValue("?", Global_Variables.random.acc_no)
                        cmd7.Parameters.AddWithValue("?", Global_Variables.random.user_name)
                        cmd7.Parameters.AddWithValue("?", Global_Variables.random.user_id)
                        cmd7.Parameters.AddWithValue("?", String.Format("Withdraw"))
                        cmd7.Parameters.AddWithValue("?", TextBox2.Text)
                        cmd7.Parameters.AddWithValue("?", String.Format("{0:dd/MM/yyyy HH:mm}", DateTime.Now))
                        cmd7.ExecuteNonQuery()
                        conn2.Close()
                    Else
                        MessageBox.Show("You dont have enough money left in the bank to withdraw any money", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    End If
                End If
            Else
                MessageBox.Show("Please Enter Valid Account No", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                TextBox1.Text = ""
                TextBox2.Text = ""
                ComboBox1.ResetText()
                DataGridView2.DataSource = ""
                DataGridView3.DataSource = ""
                DataGridView1.DataSource = ""
            End If
        End If
    End Sub
End Class