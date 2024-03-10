using MiniExcelLibs;

using FileStream fileStream = File.OpenRead(@".\issue.xlsx");
using StreamReader streamReader = new(fileStream);

var rows = MiniExcel.Query(streamReader.BaseStream, useHeaderRow: true);

for (int i = 0; i < rows.Count(); i++)
{
    Console.WriteLine($"================= ROW {i} =================");

    var row = rows.ElementAt(i);

	foreach (var cell in row)
	{
        Console.WriteLine(cell);
    }
}

