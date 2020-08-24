using System;
using System.Collections.Generic;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.Data;
using System.Reflection;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel; 


namespace Censo_DotNet_3_5
{
    class Program
    {
        private const string URL = "http://intranethspm:5001/hspmsgh-api/censo/";
        

        public static System.Data.DataTable CreateDataTable(List<Censo> arr)
        {
            XmlSerializer serializer = new XmlSerializer(arr.GetType());
            System.IO.StringWriter sw = new System.IO.StringWriter();
            serializer.Serialize(sw, arr);
            System.Data.DataSet ds = new System.Data.DataSet();
            System.Data.DataTable dt = new System.Data.DataTable();
            System.IO.StringReader reader = new System.IO.StringReader(sw.ToString());

            ds.ReadXml(reader);
            return ds.Tables[0];
        }
        static void Main(string[] args)
        {
            object misValue = System.Reflection.Missing.Value;
             DateTime today = DateTime.Now;

            System.Data.DataTable dataCenso = new System.Data.DataTable();

            List<Censo> censos = new List<Censo>();


            WebRequest request = WebRequest.Create(URL);
            try
            {
                using (var twitpicResponse = (HttpWebResponse)request.GetResponse())
                {
                    //using (var reader = new StreamReader(stream: twitpicResponse.GetResponseStream()))
                    using (var reader = new StreamReader(twitpicResponse.GetResponseStream()))
                    {
                        JsonSerializer json = new JsonSerializer();
                        var objText = reader.ReadToEnd();
                        censos = JsonConvert.DeserializeObject<List<Censo>>(objText);
                        dataCenso = CreateDataTable(censos);

                    }
                }
            }
           
            catch (Exception ex)
            {
                String error = ex.Message;
                Console.ReadKey();
            }

            String excelFilePath = "\\\\hspmins2\\NIR_Nucleo_Interno_Regulacao\\2359\\Censo" + today.ToString().Replace('/', '_').Replace(' ', '_').Replace(':', '_');
            try
            {
                if (dataCenso == null || dataCenso.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excelApp = new Excel.Application();
                Excel.Workbook workBook = excelApp.Workbooks.Add(Type.Missing);

                // single worksheet
               Microsoft.Office.Interop.Excel._Worksheet workSheet = (Microsoft.Office.Interop.Excel._Worksheet)excelApp.ActiveSheet;

                // column headings
                for (var i = 0; i < dataCenso.Columns.Count; i++)
                {
                    workSheet.Cells[1, i + 1] = dataCenso.Columns[i].ColumnName;
                }

                 //rows
                for (var i = 0; i < dataCenso.Rows.Count; i++)
                {
                    // to do: format datetime values before printing
                    for (var j = 0; j < dataCenso.Columns.Count; j++)
                    {
                       /* if (j==9 || j == 10 || j == 13 || j == 24 || j == 25  )
                        {
                            var dt = dataCenso.Rows[i][j];
                            workSheet.Cells[i + 2, j + 1] = Convert.ToDateTime(dataCenso.Rows[i][j]);
                          
                        }
                        else
                        {*/
                            workSheet.Cells[i + 2, j + 1] = dataCenso.Rows[i][j];
                           

                       // }
                    }
                }

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        
                        //workSheet.Name = "Censo" + today.ToString().Replace('/', '_');
                        workSheet.Name = "Censo";
                        workBook.SaveAs(excelFilePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                       
                        excelApp.Quit();
                        Console.WriteLine("Excel file saved!");
                      
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n" + ex.Message);
                       
                    }
                }
                else
                { // no file path is given
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception( "ExportToExcel: \n" + GetMessage(ex));

            }
            Console.ReadKey();

        }

        private static string GetMessage(Exception ex)
        {
            return ex.Message;
        }
        
    }
    public class Censo
    {


        public string Cd_prontuario { get; set; }

        public string Nm_paciente { get; set; }

        public string In_sexo { get; set; }
        public string Nr_idade { get; set; }

        public string Nr_quarto { get; set; }

        public string Dt_internacao_data { get; set; }
        public string Dt_internacao_hora { get; set; }
        public string Dt_ultimo_evento_data { get; set; }
        public string Dt_ultimo_evento_hora { get; set; }

        public string Nascimento { get; set; }
        public string Nm_unidade_funcional { get; set; }

        public string Nm_especialidade { get; set; }

        public string Nm_medico { get; set; }


        public string Nm_origem { get; set; }

        public string Sg_cid { get; set; }
        public string Descricao_cid { get; set; }


        public string Vinculo { get; set; }

        public string Nr_convenio { get; set; }

        public string Tempo { get; set; }
        public string Dt_saida_paciente { get; set; }


        internal static IEnumerable<PropertyInfo> GetProperties()
        {
            throw new NotImplementedException();
        }
    }

}
