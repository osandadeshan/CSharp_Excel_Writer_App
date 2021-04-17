namespace ExcelWriterApp
{
    class Program
    {
        static void Main()
        {
            new BasicExcelWriter().Write();
            new StylesEmbeddedExcelWriter().Write();
        }
    }
}