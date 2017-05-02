using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmoothInvoiceMailer
{
    public static class InvoiceHelper
    {
        public static string CreateInvoiceNumber()
        {
            //example 1727APR
            return DateTime.Now.Year.ToString().Substring(2, 2) + DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.ToString("MMMM").ToUpper();
        }
    }
}
