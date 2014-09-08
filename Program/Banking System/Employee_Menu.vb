
Public Class Employee_Menu

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Create_Account.Show()
    End Sub

    Private Sub HowToCreateAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToCreateAccountToolStripMenuItem.Click
        MessageBox.Show(" Use this option whenever a new customer wants to joins the bank.", "Create Account ?", MessageBoxButtons.OK)
    End Sub

    Private Sub HowToDeleteAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HowToDeleteAccountToolStripMenuItem.Click
        MessageBox.Show(" Use this option whenever a customer wants to close his/her account.", "Close Account ?", MessageBoxButtons.OK)
    End Sub

    Private Sub WhatDoesTransactionsForToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WhatDoesTransactionsForToolStripMenuItem.Click
        MessageBox.Show(" Use this option whenever a customer wants to either withdraw or deposit cash.", "Transactions ?", MessageBoxButtons.OK)
    End Sub

    Private Sub WhatToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WhatToolStripMenuItem.Click
        MessageBox.Show(" Use this option to check the details of the customer", "Employee Details ?", MessageBoxButtons.OK)
    End Sub

    Private Sub LoanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoanToolStripMenuItem.Click
        MessageBox.Show(" Use this option whenever a customer requests for a loan", "Issue Loan ?", MessageBoxButtons.OK)
    End Sub

    Private Sub ToolStripLabel2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripLabel2.Click
        Dim res As DialogResult
        Dim SAPI
        res = MessageBox.Show("Are You Sure You Want To Log Out ?", "Log Out", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If res = Windows.Forms.DialogResult.Yes Then
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

    Private Sub Employee_Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim log_in_details As String = String.Format("{0:dd/MM/yyyy" + System.Environment.NewLine + " HH:mm}", DateTime.Now)
        Label6.Text = log_in_details
        Global_Variables.random.log_in_datetime = log_in_details
        Label2.Text = Global_Variables.random.user_name + " !!!"
        Label7.Text = Global_Variables.random.user_id
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Hide()
        Employee_Details.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Close_Account.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Hide()
        Transactions.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Issue_Loan.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Hide()
        Customer_Details.Show()
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Hide()
        Clear_Loan.Show()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Me.Hide()
        Employee_Log_In_History.Show()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.Hide()
        Recent_Activity.Show()
    End Sub

    Private Sub ClearLoanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearLoanToolStripMenuItem.Click
        MessageBox.Show(" Use this option whenever a customer wants to clear his/her loan status", "Clear Loan ?", MessageBoxButtons.OK)
    End Sub

    Private Sub LogInHistoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInHistoryToolStripMenuItem.Click
        MessageBox.Show(" Select this option to check your log in details", "Log In History ?", MessageBoxButtons.OK)
    End Sub

    Private Sub MyRecentActivityToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyRecentActivityToolStripMenuItem.Click
        MessageBox.Show(" Select this option to check your recent activites", "My Recent Activity ?", MessageBoxButtons.OK)
    End Sub

    Private Sub MyDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyDetailsToolStripMenuItem.Click
        MessageBox.Show(" Select this option to view and edit your details", "My Details ?", MessageBoxButtons.OK)
    End Sub
End Class