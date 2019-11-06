using Microsoft.Csv;

namespace NetCoreInteropIssue
{
    class Program
    {
        static void Main(string[] args)
        {
            var document = new CsvDocument("Col1", "Col2");
            document.ViewInExcel();
        }
    }
}
