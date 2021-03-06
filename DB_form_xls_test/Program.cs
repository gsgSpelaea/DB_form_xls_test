﻿using System.Data;
using System.Linq;
using System;

namespace DB_form_xls_test
{
    class Program
    {
        static void Main(string[] args)
        {
            //string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location + "DataToCS.xlsx";
            string appDir = "u:\\dev\\ANYC\\DB_form_xls_test\\";
            string appPath =appDir + "DataToCS.xlsx";
            //Console.WriteLine(appPath);

            GetFromXls GetXls = new GetFromXls(appPath);

            DataSet Ds = GetXls.GetDataFromXls();

            Console.WriteLine($"\nDataTable.Count: {Ds.Tables.Count}");


            foreach (DataTable dt  in Ds.Tables)
            {
                Console.WriteLine($"\ndt.TableName: {dt.TableName}");
                Console.WriteLine($"\ndt.Rows.Count: {dt.Rows.Count}");
                Console.WriteLine($"\ndt.Columns.Count: {dt.Columns.Count}");

                // foreach ( DataRow dr in table.Rows )
                for (int i = 0; i < dt.Rows.Count; i++) 
                {

                    string strRec = "";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        strRec +=  dt.Rows[i][j].ToString() + " ";
                    }
                    Console.WriteLine(strRec);
                }
            }
            Console.ReadKey();

        }
    }
}
