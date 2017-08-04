Imports ParaLibrary.ParaLibrary
Imports System.Data.SqlClient
Imports System.Net.Mail

Module TestParaLibrary

    Public Sub sendEmail()
        Dim client As New SmtpClient("smtp.gmail.com", 587)
        client.UseDefaultCredentials = False
        client.EnableSsl = True
        client.Credentials = New Net.NetworkCredential("david.jimbo26@gmail.com", "Datnikka1!")
        Dim mail As New MailMessage("david.jimbo26@gmail.com", "drj2926@yahoo.com", "Test Send Email", getHTML())
        mail.IsBodyHtml = True
        client.Send(mail)
    End Sub

    Public Function getHTML() As String
        Dim htmlString As String = "
<table>
  <tr>
    <th>Company</th>
    <th>Contact</th>
    <th>Country</th>
  </tr>
  <tr>
    <td>Alfreds Futterkiste</td>
    <td>Maria Anders</td>
    <td>Germany</td>
  </tr>
  <tr>
    <td>Centro comercial Moctezuma</td>
    <td>Francisco Chang</td>
    <td>Mexico</td>
  </tr>
  <tr>
    <td>Ernst Handel</td>
    <td>Roland Mendel</td>
    <td>Austria</td>
  </tr>
  <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
  <tr>
    <td>Laughing Bacchus Winecellars</td>
    <td>Yoshi Tannamuri</td>
    <td>Canada</td>
  </tr>
  <tr>
    <td>Magazzini Alimentari Riuniti</td>
    <td>Giovanni Rovelli</td>
    <td>Italy</td>
  </tr>
</table>

</body>
</html>"
        Return htmlString
    End Function

    Sub Main()

        ''LOG TO DATABASE TEST
        ''Dim ex As New Exception()
        ''ex.Source = "test"
        ''logErrorsToDatabase(ex)

        ''Server = MTA2BWYPARAPP14\SQLEXPRESS; Database = ParaAuditTrail; UID = sa; pwd = 123456789"
        Dim data As New DataTable()
        Dim server As String = "DESKTOP-H5FMBDG\WEBDB"
        Dim databaseNonQuery As String = "ParaAuditTrail"
        Dim databaseQuery As String = "ParaAuditTrail"
        Dim uid As String = "sa"
        Dim pwd As String = "localadmin"
        Dim connectionNonQuery As New SqlConnection()
        connectionNonQuery = getConnection(server, databaseNonQuery, uid, pwd)
        Dim connectionQuery As New SqlConnection()
        connectionQuery = getConnection(server, databaseQuery, uid, pwd)

        sendEmail() '''
        ''TEST QUERY
        ''data = executeQuery("select * from errorlog ", connectionQuery)
        ''For Each result As DataRow In data.Rows
        ''Console.WriteLine(result("ErrorSource"))
        ''Next


        ''TEST NonQUERY
        ''executeNonQuery("Insert into errorlog values('I','am','thebest')", connectionNonQuery)

        ''TEST EMAIL
    End Sub

End Module