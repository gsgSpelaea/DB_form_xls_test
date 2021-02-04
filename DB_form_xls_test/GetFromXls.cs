using System;
using System.Collections.Generic;
using System.Text;
using ExcelDataReader;
using System.IO;
using System.Data;


namespace DB_form_xls_test
{
    public class GetFromXls
    {

        public string FilePath { get; set; }


        public GetFromXls(string filePath) { FilePath = filePath; }
    //public filepath {get;set;};



    public DataSet GetDataFromXls()//(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            DataSet result = null;

            using (var stream = File.Open(this.FilePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream,
                    new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252)

                    }))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                     result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                }
            }

            return result;
        }

    }
}
