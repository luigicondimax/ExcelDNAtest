using ExcelDna.Integration;

namespace ExcelDNAtest
{
    public class Function
    {
        [ExcelFunction]
        public static string Hello()
        {
            return "Hello world!";
        }
    }
}
