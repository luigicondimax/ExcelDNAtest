using ExcelDna.Integration;

namespace ExcelDNAtest
{
    public class Funzioni
    {
        [ExcelFunction]
        public static string Hello()
        {
            return "Hello world!";
        }
    }
}
