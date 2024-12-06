using static System.Environment;
using System.Data;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
        string connectionString = "Server=BYWPWDB06.byl.com;Database=ER;User Id=ER2AgentWeb;Password=UG83Ic&XRfV5;";
        string storedProcedureName = "[dbo].[FindBadAccounts]";
        string downloadsPath = Path.Combine(GetFolderPath(SpecialFolder.UserProfile), "Downloads");
        string filePath = Path.Combine(downloadsPath, "output.xlsx");

        // Step 1: Get the matching IDs from the query
        var matchingIds = GetMatchingIds(connectionString);

        // Step 2: Get stored procedure results
        DataTable spResults = GetStoredProcedureResults(connectionString, storedProcedureName);

        // Step 3: Filter results and include only rows where AccountId matches
        DataTable filteredResults = new DataTable();
        filteredResults.Columns.Add("AccountId", typeof(int)); // Add AccountId column
        filteredResults.Columns.Add("Message", typeof(string)); // Add renamed column 'Message' (formerly 'MismatchReason')

        foreach (DataRow row in spResults.Rows)
        {
            if (row["AccountId"] != DBNull.Value && matchingIds.Contains((int)row["AccountId"]))
            {
                // Create a new row for the filtered results containing only AccountId and Message
                DataRow newRow = filteredResults.NewRow();
                newRow["AccountId"] = row["AccountId"];
                newRow["Message"] = row["MismatchReason"]; // Copy MismatchReason to Message column
                filteredResults.Rows.Add(newRow);
            }
        }

        // Step 4: Check if there are any rows to export
        if (filteredResults.Rows.Count > 0)
        {
            // Step 5: Export the filtered results to Excel
            ExportToExcel(filteredResults, filePath);
            Console.WriteLine($"Excel exported successfully to: {filePath}");
        }
        else
        {
            Console.WriteLine("No data to export.");
            Console.ReadLine();
        }        
    }

    // Step 1: Get the matching AccountIds based on the query
    static HashSet<int> GetMatchingIds(string connectionString)
    {
        string query = @"SELECT distinct ClientCustomerAccountID from ARTrx where TrxTypeID<>1 and cast(CreatedOn as date)=cast(getdate()-1 as date)";

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
}