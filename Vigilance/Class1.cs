using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Vigilance
{
    public class Condition
    {
        public Relation relation { get; set; }
        public Microsoft.Office.Interop.Excel.Range range { get; set; }
        public double starndard { get; set; }
    }

    public enum Relation
    {
        GreaterthanorEqualto,
        Greaterthan,
        Equalto,
        Lessthan,
        LessthanorEqualto
    }
}
