using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ValutaUpdater
{
    class Program
    {
        static void Main(string[] args)
        {
            DBFWork v2dbf = new DBFWork("V2.DBF");
            v2dbf.InsertNew(10,"FFF",1.2345,"998",DateTime.Now);
        }
    }
}
