using ExcelDna.Integration;

namespace ExcelDNAtest
{
    public class Function
    {
        [ExcelFunction(description:"Descrizione")]
        public static string Hello()
        {
            return "Hello world!";
        }

    }
}
