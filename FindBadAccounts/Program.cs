﻿using static System.Environment;
using System.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Net.Mail;
using System.Net;
using Microsoft.Extensions.Configuration;

class Program
{
    static void Main()
    {
        // Load configuration
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // Fetch configuration values
        string connectionString = config.GetConnectionString("DefaultConnection");
        string storedProcedureName1 = config["StoredProcedures:FindBadAccounts"];
        string storedProcedureName2 = config["StoredProcedures:FixZeroBalanceAccountStatus"];
        var mailSettings = config.GetSection("EmailSettings").Get<MailSettings>();
        int TrxDay = Convert.ToInt32(config["TransactionDay"]);

        string downloadsPath = Path.Combine(GetFolderPath(SpecialFolder.UserProfile), "Downloads");
        string filePath = Path.Combine(downloadsPath, "output.xlsx");

        // Step 1: Get the matching IDs from the query and update account status if incorrect
        var matchingIds = GetMatchingIds(connectionString, TrxDay);
        CorrectStatus(matchingIds, connectionString, storedProcedureName2);

        // Step 2: Get stored procedure results
        DataTable spResults = GetStoredProcedureResults(connectionString, storedProcedureName1);

        // Step 3: Filter results and include only rows where AccountId matches
        DataTable filteredResults = new DataTable();
        filteredResults.Columns.Add("AccountId", typeof(int)); // Add AccountId column
        filteredResults.Columns.Add("Message", typeof(string)); // Add renamed column 'Message'

        // Create a lookup dictionary from spResults for fast ID lookup
        var spResultsLookup = spResults.AsEnumerable()
            .Where(row => row["AccountId"] != DBNull.Value)
            .ToDictionary(row => (int)row["AccountId"], row => row["MismatchReason"]?.ToString());

        // Iterate through matchingIds (smaller collection)
        foreach (int id in matchingIds)
        {
            if (spResultsLookup.TryGetValue(id, out string message))
            {
                // Create a new row for the filtered results
                DataRow newRow = filteredResults.NewRow();
                newRow["AccountId"] = id;
                newRow["Message"] = message;
                filteredResults.Rows.Add(newRow);
            }
        }

        // Step 4: Check if there are any rows to export
        if (filteredResults.Rows.Count > 0)
        {
            // Step 5: Export the filtered results to Excel
            ExportToExcel(filteredResults, filePath);
            SendEmail(true, mailSettings, filePath);
            File.Delete(filePath);
        }
        else
        {
            SendEmail(false, mailSettings);
            Console.WriteLine("No data to export.");
        }
        Console.ReadLine();
    }

    // Step 1: Get the matching AccountIds based on the query
    static HashSet<int> GetMatchingIds(string connectionString, int TrxDay)
    {
        string query = $"SELECT distinct ClientCustomerAccountID from ARTrx where TrxTypeID<>1 and cast(CreatedOn as date)=cast(getdate()-{TrxDay} as date)";

        HashSet<int> ids = new HashSet<int>();

        using (SqlConnection connection = new SqlConnection(connectionString))
        using (SqlCommand command = new SqlCommand(query, connection))
        {
            connection.Open();
            using (SqlDataReader reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    ids.Add(reader.GetInt32(0)); // Assuming 'id' is an integer
                }
            }
        }
        return ids;
    }

    // Step 2: Get results from stored procedure
    static DataTable GetStoredProcedureResults(string connectionString, string storedProcedureName)
    {
        DataTable dataTable = new DataTable();

        using (SqlConnection connection = new SqlConnection(connectionString))
        using (SqlCommand command = new SqlCommand(storedProcedureName, connection))
        {
            command.CommandType = CommandType.StoredProcedure;

            // Increase timeout to allow longer execution
            command.CommandTimeout = 300; // Timeout in seconds (set as needed)

            // Add parameters to the command
            command.Parameters.AddWithValue("@clientCustomerAccountId", 0);
            command.Parameters.AddWithValue("@checkForStatusIssue", 0);

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            adapter.Fill(dataTable);
        }

        return dataTable;
    }


    // Step 4: Export the filtered DataTable to Excel
    static void ExportToExcel(DataTable dataTable, string filePath)
    {
        using XLWorkbook workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        // Add DataTable headers
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            worksheet.Cell(1, col + 1).Value = dataTable.Columns[col].ColumnName;
        }

        // Add DataTable rows
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                worksheet.Cell(row + 2, col + 1).Value = dataTable.Rows[row][col]?.ToString();
            }
        }

        // Auto-fit columns for better readability
        worksheet.Columns().AdjustToContents();

        // Save the file
        workbook.SaveAs(filePath);
    }

    static void CorrectStatus(HashSet<int> Ids, string connectionString, string storedProcedureName)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open(); // Open the connection once to optimize performance

            foreach (int id in Ids)
            {
                using (SqlCommand command = new SqlCommand(storedProcedureName, connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@ClientCustomerAccountID", id);

                    try
                    {
                        command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        // Log or handle the exception as needed, or suppress if required
                        Console.WriteLine($"Error executing stored procedure \"{storedProcedureName}\" for ID {id}: {ex.Message}");
                    }
                }
            }
        }
    }


    static void SendEmail(bool HasData, MailSettings mailSettings, string attachmentFilePath = null)
    {
        try
        {
            // SMTP Configuration
            string smtpServer = mailSettings.SmtpServer;
            int smtpPort = mailSettings.SmtpPort; // Default port for STARTTLS
            string fromEmail = mailSettings.FromEmail;
            string emailPassword = mailSettings.EmailPassword;
            string body = "";
            string content = "";

            content = HasData ? "Please find attached ER bad accounts." : "No bad account found.";

            body = $"<!DOCTYPE html><html><head><style>body{{font-family:Arial,sans-serif;font-size:14px;color:#333333;line-height:1.6;}}.email-container{{margin:0;padding:10px;border:1px solid #dddddd;border-radius:5px;background-color:#f9f9f9;}}.header{{font-weight:bold;margin-bottom:10px;}}.content{{margin-bottom:10px;}}</style></head><body><div class='email-container'><div class='header'>Hi,</div><div class='content'>{content}</div></div></body></html>";


            // Create a MailMessage object
            MailMessage mail = new MailMessage
            {
                From = new MailAddress(fromEmail),
                Subject = "ER Bad Account",
                Body = body,
                IsBodyHtml = true // Set to true if the email body contains HTML
            };

            // Add recipient
            mail.To.Add(mailSettings.ToEmail);

            // Add attachment if a file path is provided
            if (!string.IsNullOrEmpty(attachmentFilePath) && File.Exists(attachmentFilePath))
            {
                Attachment attachment = new Attachment(attachmentFilePath);
                mail.Attachments.Add(attachment);
            }
            else if (!string.IsNullOrEmpty(attachmentFilePath))
            {
                Console.WriteLine("Attachment file not found. Skipping attachment.");
            }

            // Configure the SMTP client
            SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort)
            {
                Credentials = new NetworkCredential(fromEmail, emailPassword),
                EnableSsl = true // STARTTLS requires SSL
            };

            // Send the email
            smtpClient.Send(mail);
            Console.WriteLine("Email sent successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to send email. Error: {ex.Message}");
        }
    }
}

// Class for mail settings
public class MailSettings
{
    public string SmtpServer { get; set; }
    public int SmtpPort { get; set; }
    public string FromEmail { get; set; }
    public string EmailPassword { get; set; }
    public string ToEmail { get; set; }
}