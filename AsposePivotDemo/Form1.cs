using System;

using System.Collections.Generic;

using System.ComponentModel;

using System.Data;

using System.Drawing;

using System.Linq;

using System.Text;

using System.Threading.Tasks;

using System.Windows.Forms;

using Aspose.Cells;

using Aspose.Cells.Charts;

using Aspose.Cells.Pivot;

using Newtonsoft.Json;

using System.IO;

using Newtonsoft.Json.Linq;



namespace AsposePivotDemo

{

    public partial class Form1 : Form

    {

        public Form1()

        {

            InitializeComponent();

        }



        private void button1_Click(object sender, EventArgs e)

        {

            Workbook workbook = new Workbook("C:/Users/Ankita/Downloads/Demo2.xlsx");

            //workbook.Worksheets.RemoveAt("Pivot");

            Worksheet worksheet = workbook.Worksheets.Add("Pivot");





            #region removing hard coding

            Worksheet sheet0 = workbook.Worksheets[0];

            Cells all_cells = sheet0.Cells;



            Cell cell = sheet0.Cells.LastCell;

            Cell cell_first = sheet0.Cells.FirstCell;



            #region Last cell calculation





            int col_last = cell.Column + 1;

            int row_last = cell.Row + 1;



            #endregion



            #region First calculation



            int col_first = cell_first.Column + 1;

            int row_first = cell_first.Row + 1;



            #endregion

















            #region Calculate column character

            int dividend = col_last;

            string columnName = String.Empty;



            while (dividend > 0)

            {

                var modulo = (dividend - 1) % 26;

                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;

                dividend = (dividend - modulo) / 26;

            }



            #endregion

            //To read the sheet name

            string sheetname = sheet0.Name;

            #endregion

            string datasource = sheetname + "!A1:" + columnName + row_last.ToString();



            //string datasource = sheetname + "!A" + row_first.ToString() + ":" + columnName + row_last.ToString();



            //JObject o1 = JObject.Parse(File.ReadAllText("C:/Users/Ankita.Gopakumar/Downloads/File.json"))

            //JArray spendUSD = (JArray)o1["SpendUSD"];

            //Console.WriteLine(spendUSD);

            #region JSON

            // using (StreamReader r = new StreamReader(@"C:\Users\Ankita\Source\Repos\AsposePivotDemoSolution-master\AsposePivotDemo\File.json"))

            // {

            //     string json = r.ReadToEnd();

            //     dynamic array = JsonConvert.DeserializeObject(json);

            //     var result = new Dictionary<string, string>();



            //     foreach (var field in array)

            //     {

            //        result.Add(field.CFDCOL, Convert.ToString(field.FieldValue.value));

            //     }

            //     foreach (var item in result)

            //     {

            //        Console.WriteLine(item.Key + "" + item.Value);

            //    }

            // }





            #endregion







            int iPivotIndex = worksheet.PivotTables.Add(datasource, "A1", "PivotTable");

            PivotTable pt = worksheet.PivotTables[iPivotIndex];

            pt.RowGrand = false;

            pt.ColumnGrand = false;

            pt.IsAutoFormat = true;

            pt.AddFieldToArea(PivotFieldType.Row, 0);

            pt.AddFieldToArea(PivotFieldType.Row, 1);

            pt.AddFieldToArea(PivotFieldType.Data, 2);

            #region Globalisation

            Cell cell1 = sheet0.Cells["C1"];

            Style style = sheet0.Cells["C1"].GetStyle();



            //string currencySymbol = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol.ToString();



            //style.Custom = "£#,##0;[Red]-£#,##0";

            //style.Number = 5;

            string value = cell1.StringValue;

            if (value == "Spend USD")

            {

                //worksheet.Cells["C1"].SetStyle(style);

                pt.DataFields[0].NumberFormat = @"[>=1000000]$###\,###\,##0.00;[>=100000] $###\,##0;##,##0.00";

            }

            else

            {

                pt.DataFields[0].NumberFormat = @"[>=1000000]€###\,###\,##0.00;[>=100000] €###\,##0;##,##0.00";

                // pt.DataFields[0].NumberFormat = @"[>=1000000]€###\.###\.##0\,00;[>=100000] €###\.##0;##.##0\,00";

                //style.Custom = "€#.##0.00_);[Red](€#.##0.00)";

                // StyleFlag flg = new StyleFlag();

                //flg.NumberFormat = true;

                //string custom = style.Custom;



                //string cultureCustom = style.CultureCustom;

                //Style newStyle = workbook.CreateStyle();



                //newStyle.CultureCustom = g;



                //newStyle.Custom = “#,##0.000\ [$€-40C]”;

                //cell1.SetStyle(newStyle);

                //pt.DataFields[0].Number=4;

            }

            #endregion

            pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleDark1;



            //workbook.Worksheets[0].IsVisible = false;



            Style st = workbook.CreateStyle();

            pt.FormatAll(st);





            workbook.Save("C:/Users/Ankita/Downloads/Demo2.xlsx");


            //trial



           


        }

    }

}