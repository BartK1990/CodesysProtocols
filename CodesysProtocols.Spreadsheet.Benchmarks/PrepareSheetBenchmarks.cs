using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using CodesysProtocols.Spreadsheet.ExcelAccess;
using OfficeIMO.Excel;

namespace CodesysProtocols.Spreadsheet.Benchmarks;

[SimpleJob(
    RunStrategy.Throughput, 
    launchCount: 1,
    warmupCount: 2,
    iterationCount: 3,
    invocationCount: 2)]
[MemoryDiagnoser]
//[EtwProfiler]
public class PrepareSheetBenchmarks
{
    private ExcelDocument _doc = default!;
    private ExcelSheet _sheet = default!;
    private string _tempFile = default!;

    [Params(50, 200, 500)]
    public int ColumnCount;

    [GlobalSetup]
    public void Setup()
    {
        _tempFile = Path.ChangeExtension(Path.GetTempFileName(), ".xlsx");
        _doc = ExcelDocument.Create(_tempFile, "Sheet1");
        _sheet = _doc.Sheets[0];

        // Create a 3-row header with many columns.
        for (int c = 1; c <= ColumnCount; c++)
        {
            _sheet.CellValue(1, c, "Node");
            _sheet.CellValue(2, c, "Name");
            _sheet.CellValue(3, c, "Type");
        }
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        _doc.Dispose();
        if (!string.IsNullOrWhiteSpace(_tempFile) && File.Exists(_tempFile))
        {
            File.Delete(_tempFile);
        }
    }

    [Benchmark]
    public void PrepareSheet()
    {
        // Write a minimal table so that it calls PrepareSheet.
        var table = new Model.TableData.Iec608705.Iec608705Table { Name = "Sheet1" };
        for (int c = 0; c < ColumnCount; c++)
        {
            table.Columns.Add(new Model.TableData.Iec608705.Iec608705Column
            {
                NodeName = "Node",
                Name = "Name",
                Type = "Type",
            });
        }

        ExcelIec608705Table.Write(table, _sheet);
    }
}
