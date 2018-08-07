﻿using System;
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
            Workbook workbook = new Workbook("C:/Users/Ankita.Gopakumar/Downloads/Demo.xlsx");
            //workbook.Worksheets.RemoveAt("Pivot");
            Worksheet worksheet = workbook.Worksheets.Add("Pivot");


            #region removing hard coding 
            Worksheet sheet0 = workbook.Worksheets[0];
            //Cells cells = sheet0.Cells;

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
            using (StreamReader r = new StreamReader("C:/Users/Ankita.Gopakumar/Downloads/File.json"))
            {
                string json = r.ReadToEnd();
                dynamic array = JsonConvert.DeserializeObject(json);
                var result = new Dictionary<string, string>();

                foreach (var field in array)
                {
                    result.Add(field.Name, Convert.ToString(field.V));
                }
                foreach (var item in result)
                {
                    Console.WriteLine(item.Key+""+item.Value);
                }
            }

            int iPivotIndex = worksheet.PivotTables.Add(datasource, "A1", "PivotTable");
            PivotTable pt = worksheet.PivotTables[iPivotIndex];
            pt.RowGrand = false;
            pt.ColumnGrand = false;
            pt.IsAutoFormat = true;
            pt.AddFieldToArea(PivotFieldType.Row, 0);
            pt.AddFieldToArea(PivotFieldType.Row, 1);
            pt.AddFieldToArea(PivotFieldType.Data, 2);

            pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleDark1;

            //workbook.Worksheets[0].IsVisible = false;

            Style st = workbook.CreateStyle();
            pt.FormatAll(st);

            workbook.Save("C:/Users/Ankita.Gopakumar/Downloads/Demo.xlsx");




        }
    }
}