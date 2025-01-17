using static System.Environment;
using System.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using System.Net.Mail;
using System.Net;
using Microsoft.Extensions.Configuration;
using System.Globalization;

class Program
{
    static async Task Main()
    {
        try
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

            // Step 2: Get stored procedure results
            DataTable spResults = await GetStoredProcedureResultsAsync(connectionString, storedProcedureName1, TrxDay);

            // Step 3: Filter results and include only rows where AccountId matches
            DataTable filteredResults = new DataTable();
            filteredResults.Columns.Add("AccountId", typeof(int)); // Add AccountId column
            filteredResults.Columns.Add("Message", typeof(string)); // Add renamed column 'Message'

            // Loop through each row in spResults
            foreach (DataRow row in spResults.Rows)
            {
                // Extract the AccountId and Message values
                DataRow newRow = filteredResults.NewRow();
                newRow["AccountId"] = Convert.ToInt32(row["AccountId"]);
                newRow["Message"] = row["MismatchReason"]?.ToString();
                filteredResults.Rows.Add(newRow);
            }

            if (filteredResults.Rows.Count > 0)
            {
                await CorrectStatusAsync(filteredResults, connectionString, storedProcedureName2);
            }

            // Step 4: Check if there are any rows to export
            if (filteredResults.Rows.Count > 0)
            {
                // Step 5: Export the filtered results to Excel
                await ExportToExcelAsync(filteredResults, filePath);
                await SendEmailAsync(true, mailSettings, TrxDay, filePath);
                File.Delete(filePath);
            }
            else
            {
                await SendEmailAsync(false, mailSettings, TrxDay);
                Console.WriteLine("No data to export.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // Step 4: Export the filtered DataTable to Excel
    static async Task ExportToExcelAsync(DataTable dataTable, string filePath)
    {
        await Task.Run(() =>
        {
            using XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // Add DataTable headers
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                worksheet.Cell(1, col + 1).Value = dataTable.Columns[col].ColumnName;
                worksheet.Cell(1, col + 1).Style.Font.FontName = "Roboto"; // Set font type
                worksheet.Cell(1, col + 1).Style.Font.FontSize = 10;       // Set font size
            }

            // Add DataTable rows
            for (int row = 0; row < dataTable.Rows.Count; row++)
            {
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    var cell = worksheet.Cell(row + 2, col + 1);
                    cell.Value = dataTable.Rows[row][col]?.ToString();
                    cell.Style.Font.FontName = "Roboto"; // Set font type
                    cell.Style.Font.FontSize = 10;       // Set font size
                }
            }

            // Apply borders to all cells with data
            var range = worksheet.RangeUsed();
            if (range != null)
            {
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            // Auto-fit columns for better readability
            worksheet.Columns().AdjustToContents();

            // Save the file
            workbook.SaveAs(filePath);
        });
    }


    static async Task<DataTable> GetStoredProcedureResultsAsync(string connectionString, string storedProcedureName, int TrxDay)
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
            command.Parameters.AddWithValue("@transactionyear", 0);
            command.Parameters.AddWithValue("@trxday", TrxDay);

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            await Task.Run(() => adapter.Fill(dataTable));
        }

        return dataTable;
    }

    static async Task CorrectStatusAsync(DataTable Ids, string connectionString, string storedProcedureName)
    {
        using SqlConnection connection = new SqlConnection(connectionString);
        await connection.OpenAsync(); // Open the connection once to optimize performance

        foreach (DataRow row in Ids.Rows)
        {
            using SqlCommand command = new SqlCommand(storedProcedureName, connection);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@ClientCustomerAccountID", row["AccountId"]);

            try
            {
                await command.ExecuteNonQueryAsync();
            }
            catch (Exception ex)
            {
                // Log or handle the exception as needed, or suppress if required
                Console.WriteLine($"Error executing stored procedure \"{storedProcedureName}\" for ID {row["AccountId"]}: {ex.Message}");
            }
        }
    }

    static async Task SendEmailAsync(bool HasData, MailSettings mailSettings, int TrxDay, string attachmentFilePath = null)
    {
        try
        {
            // SMTP Configuration
            string smtpServer = mailSettings.SmtpServer;
            int smtpPort = mailSettings.SmtpPort; // Default port for STARTTLS
            string fromEmail = mailSettings.FromEmail;
            string emailPassword = mailSettings.EmailPassword;
            string content = HasData ? "Please find attached ER bad accounts." : "No bad account found.";

            string body = $"<!DOCTYPE html><html><head><style>body{{font-family:Arial,sans-serif;font-size:14px;color:#333333;line-height:1.6;}}.email-container{{margin:0;padding:10px;border:1px solid #dddddd;border-radius:5px;background-color:#f9f9f9;}}.header{{font-weight:bold;margin-bottom:10px;}}.content{{margin-bottom:10px;}}</style></head><body><div class='email-container'><div class='header'>Hi,</div><div class='content'>{content}</div></div></body></html>";

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(fromEmail);

                TimeZoneInfo estZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                DateTime estTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, estZone);
                mail.Subject = $"ER Bad Account : {estTime.AddDays(-1 * TrxDay).Date.ToString("dd-MM-yyyy", new CultureInfo("hi-IN"))}";

                mail.Body = body;
                mail.IsBodyHtml = true; // Set to true if the email body contains HTML

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
                using (SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort))
                {
                    smtpClient.Credentials = new NetworkCredential(fromEmail, emailPassword);
                    smtpClient.EnableSsl = true; // STARTTLS requires SSL

                    // Send the email
                    await smtpClient.SendMailAsync(mail);
                    Console.WriteLine("Email sent successfully!");
                }
            }
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