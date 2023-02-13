// See https://aka.ms/new-console-template for more information
using ClosedXML.Report;
using Microsoft.Data.SqlClient;
using System.Data;

Console.WriteLine("Collegamento al database");


    

var template = new XLTemplate(@".\Template.xlsx");

template.Workbook.Worksheets.TryGetWorksheet("Queries", out var queriesWS);
if (queriesWS == null) { 
    throw new ArgumentNullException(nameof(queriesWS));
}

bool summaryInfoSet = false;
using (var cn = new SqlConnection(@"Server=server;Database=database;Trusted_Connection=True;Encrypt=False")) {
    cn.Open();

    var lastCell = queriesWS.LastCellUsed();
    for (var row = 2; row <= lastCell.WorksheetRow().RowNumber(); row++) {
        var queryName = queriesWS.Cell(row, 1).GetString();
        var queryStatement = queriesWS.Cell(row, 2).GetString();
        if (string.IsNullOrEmpty(queryStatement)) { continue; }

        var dt = new DataTable();
        var da = new SqlDataAdapter(queryStatement, cn);
        da.Fill(dt);

        if (string.IsNullOrEmpty(queryName)) {
            if (!summaryInfoSet) {
                template.AddVariable(dt);
                summaryInfoSet = true;
            } else {
                throw new Exception("Multiple queries as Summary Info");
            }
        } else {
            template.AddVariable(queryName, dt);
        }
    }
}
if (!summaryInfoSet) {
    // Qui si potrebbero aggiungere informazioni di contesto per il report
    var summaryInfo = new { Title = "Export con ClosedXmlReport" };
    template.AddVariable(summaryInfo);
}

template.Generate();
template.Workbook.Worksheets.Delete("Queries");

template.SaveAs(@".\..\..\..\Output.xlsx");

