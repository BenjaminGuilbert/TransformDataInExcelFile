using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace Idema_Import_UpSell
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = ReadExcelFile();

            Dictionary<int, List<String>> map = new Dictionary<int, List<string>>();


            //Stocker toutes les données
            foreach (DataRow row in dt.Rows)
            {
                int page = Int32.Parse(row[0].ToString());
                string product = row[1].ToString();

                //Console.WriteLine(""+page +":"+product);

                if (!map.ContainsKey(page))
                {
                    map.Add(page, new List<String>());
                }

                map[page].Add(product);
                
            }

            Console.WriteLine(map.Count);


            //Transformer les données
            Dictionary<string, string> result = new Dictionary<string, string>();

            foreach (KeyValuePair<int, List<String>> page in map)
            {
                //Console.WriteLine(page.Key + ":" + page.Value.ToArray().Length);
                foreach (string prod in page.Value)
                {
                    if (!result.ContainsKey(prod))
                    {
                        result.Add(prod, "");
                    }              


                    foreach (string upsell in page.Value)
                    {
                        if (!prod.Equals(upsell))
                        {
                            result[prod] += ";" + upsell;
                        }
                    }                 

                }
            }


            //Sauvegarder les données dans une nouvelle sheet
            WriteExcelFile(result);

            Console.Read();
            
        }


        /// <summary>
        /// Read data from a Excel file
        /// </summary>
        /// <returns>DatatTable</returns>
        private static DataTable ReadExcelFile()
        {
            DataTable dt = new DataTable();

            string connectionString = GetConnectionString();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                //Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                //Get the first sheet
                DataRow dr = dtSheet.Rows[0];
                string sheetName = dr["TABLE_NAME"].ToString();
                dt.TableName = sheetName;

                //Get all data in this sheet
                cmd.CommandText = "SELECT * FROM [" + sheetName + "]";                

                //Fill the table with data from the sheet
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);

                cmd = null;
                conn.Close();



            }

            return dt;
        }

        /// <summary>
        /// Write data in the file in a new sheet
        /// </summary>
        /// <param name="map">map to 2 columns</param>
        private static void WriteExcelFile(Dictionary<string, string> map)
        {
            string connectionString = GetConnectionString();

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                cmd.CommandText = "CREATE TABLE [result] (product VARCHAR, upsells VARCHAR);";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO [result](product,upsells) VALUES(@prod,@upsells);";
                foreach (KeyValuePair<string, string> prod in map)
                {
                    Console.WriteLine(prod.Key + ":" +prod.Value.Length);
                    Console.WriteLine(prod.Value);

                    String upsells = prod.Value;
                    String restUpSells = null;
                    if (upsells.Length > 255)
                    {
                        int startSecondPart = prod.Value.IndexOf(";",200);
                        upsells=prod.Value.Substring(0, startSecondPart);
                        restUpSells = prod.Value.Substring(startSecondPart,prod.Value.Length-startSecondPart);
                    }          
          
                    cmd.Parameters.AddWithValue("@prod", prod.Key);
                    cmd.Parameters.AddWithValue("@upsells", upsells);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    if (!(restUpSells == null))
                    {
                        cmd.Parameters.AddWithValue("@prod", prod.Key);
                        cmd.Parameters.AddWithValue("@upsells", restUpSells);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }

                }

                


                conn.Close();
            }
        }


        /// <summary>
        /// Get the connection to the file
        /// </summary>
        /// <returns></returns>
        private static string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            // XLSX - Excel 2007, 2010, 2012, 2013
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = @"C:\TEMP\IdemaSport\Produits_Up_sell.xls";

            // XLS - Excel 2003 and Older
            //props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
            //props["Extended Properties"] = "Excel 8.0";
            //props["Data Source"] = @"C:\TEMP\IdemaSport\Produits_Up_sell.xls";

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
    }
}
