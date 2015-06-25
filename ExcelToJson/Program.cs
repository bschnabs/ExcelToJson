using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data;
using System.Data.OleDb;
using System.Data.Entity;

namespace ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            //salesRepID = newSet.Tables[0].Select().GroupBy(x => x["Name"]).Select(grp => grp.First()).FirstOrDefault()["Name"].ToString()
            DataSet newSet = excelConversion.excelSet(@"E:\SaberLogic\Kaufman\Sales_RFQ\Customer Listing by Salesman.xlsx");

            var listOfReps = newSet.Tables[0].AsEnumerable().GroupBy(x => x["Name"]).Select(dr => new SalesReps
                                                            {
                                                                salesRepID = dr.FirstOrDefault()["Name"].ToString(),
                                                                name = dr.FirstOrDefault()["Name"].ToString(),
                                                                phoneNum = 8888888888,
                                                                emailAdd = "test@kaufman.com",
                                                                customers = newSet.Tables[0].
                                                                                AsEnumerable().
                                                                                Where(x => (string)x["Name"] == (string)dr.
                                                                                FirstOrDefault()["Name"]).Select(cr => new customers
                                                                                                          {
                                                                                                            customerID = cr["Cust Num"].ToString(),
                                                                                                            custName = (string)cr["Customer Name"]
                                                                                                          }).ToArray()
                                                            }).ToList();

            Console.ReadLine();
        }
    }

    class excelConversion
    {
        public static DataSet excelSet(string fileName)
        {
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            var adapter = new OleDbDataAdapter("select * from [kauf$]", connectionString);
            var ds = new DataSet();

            adapter.Fill(ds, "Customers");

            return ds;
        }

    }

    class SalesReps
    {
        public string salesRepID { get; set; }
        public string name { get; set; }
        public long phoneNum { get; set; }
        public string emailAdd { get; set; }
        public customers[] customers { get; set; }
    }

    class customers
    {
        public string salesRepID { get; set; }
        public string customerID { get; set; }
        public string custName { get; set; }
        public string add1 { get; set; }
        public string add2 { get; set; }
        public string add3 { get; set; }
        public string city { get; set; }
        public string state { get; set; }
        public string zip { get; set; }
        public string primaryContact { get; set; }
        public customerContact[] contacts { get; set; }
    }

    class customerContact
    {
        public string name { get; set; }
        public int phone { get; set; }
        public int fax { get; set; }
    }
}
