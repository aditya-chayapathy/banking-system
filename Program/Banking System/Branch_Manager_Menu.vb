Public Class Branch_Manager_Menu

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Branch_Manager_Menu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim log_in_details As String = String.Format("{0:dd/MM/yyyy" + System.Environment.NewLine + " HH:mm}", DateTime.Now)
        Label6.Text = log_in_details
        Global_Variables.random.log_in_datetime = log_in_details
        Label2.Text = Global_Variables.random.user_name + " !!!"
        Label7.Text = Global_Variables.random.user_id
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        Hire_Employee.Show()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Dismiss_Employee.Show()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Hide()
        List_Of_Borrowers.Show()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Hide()
        Branch_Transactions.Show()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
        Branch_Statistics.Show()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Me.Hide()
        Branch_Manager_Log_In_History.Show()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Hide()
        Branch_Manager_Details.Show()
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

    Private Sub HireEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HireEmployeeToolStripMenuItem.Click
        MessageBox.Show("Select this option to add a new employee to the bank", "Hire An Employee ?", MessageBoxButtons.OK)
    End Sub

    Private Sub DismissEmployeeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DismissEmployeeToolStripMenuItem.Click
        MessageBox.Show("Select this option to dismiss an employee from the bank", "Dismiss An Employee ?", MessageBoxButtons.OK)
    End Sub

    Private Sub ListOfBorrowersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListOfBorrowersToolStripMenuItem.Click
        MessageBox.Show("Select this option to view the list of customers who have taken loans", "List Of Borrowers ?", MessageBoxButtons.OK)
    End Sub

    Private Sub BranchTransactionsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BranchTransactionsToolStripMenuItem.Click
        MessageBox.Show("Select this option to view the transactions that have been carried out", "Bank Transactions ?", MessageBoxButtons.OK)
    End Sub

    Private Sub BankStatisticsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BankStatisticsToolStripMenuItem.Click
        MessageBox.Show("Select this option to view branch statistics", "Branch Statistics ?", MessageBoxButtons.OK)
    End Sub

    Private Sub MyRecentActivitiesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MessageBox.Show("Select this option to view your recent activities", "My Recent Activites ?", MessageBoxButtons.OK)
    End Sub

    Private Sub LogInHistoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogInHistoryToolStripMenuItem.Click
        MessageBox.Show("Select this option to view your log in history", "Log In History ?", MessageBoxButtons.OK)
    End Sub

    Private Sub MyDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyDetailsToolStripMenuItem.Click
        MessageBox.Show("Select this option to view and/or edit your details", "My Details ?", MessageBoxButtons.OK)
    End Sub

    Private Sub EmployeeDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmployeeDetailsToolStripMenuItem.Click
        MessageBox.Show("Select this option to view the details of the employees present in the bank", "Employee Details ?", MessageBoxButtons.OK)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Me.Hide()
        Employee_Details_Branch_Manager.Show()
    End Sub
End Class