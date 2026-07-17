using ExcelDna.Integration;

namespace ExcelDNAtest
{
    public class Function
    {
        [ExcelFunction(description:"Descrizione",ExplicitRegistration =true)]
        public static string Hello()
        {
            return "Hello world!";
        }

    }
}
