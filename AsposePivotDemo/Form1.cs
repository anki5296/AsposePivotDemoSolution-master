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

            string file_location = "C:/Users/Ankita/Downloads/global.xlsx";

            Workbook workbook = new Workbook(file_location);

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




            #region JSON

            //string datasource = sheetname + "!A" + row_first.ToString() + ":" + columnName + row_last.ToString();



            




            #endregion
            int iPivotIndex = worksheet.PivotTables.Add(datasource, "A1", "PivotTable");

            PivotTable pt = worksheet.PivotTables[iPivotIndex];

            pt.RowGrand = false;

            pt.ColumnGrand = false;

            pt.IsAutoFormat = true;

            pt.AddFieldToArea(PivotFieldType.Column, 0);
            pt.AddFieldToArea(PivotFieldType.Row, 1);
            pt.AddFieldToArea(PivotFieldType.Data, 2);
            pt.AddFieldToArea(PivotFieldType.Data, 3);

            #region Globalisation

            Cell cell1 = sheet0.Cells["C1"];
            Cell cell2 = sheet0.Cells["D1"];
            Style style = sheet0.Cells["C1"].GetStyle();

            //string currencySymbol = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol.ToString();
            //style.Custom = "£#,##0;[Red]-£#,##0";
            //style.Number = 5;

            string value = cell1.StringValue;
            string value1 = cell2.StringValue;

            if (value == "Spend USD")

            {

                //worksheet.Cells["C1"].SetStyle(style);

                pt.DataFields[0].NumberFormat = @"[>999999999]$#\,###\,###\,##0.00;[>999999]$###\,###\,##0.00;$#,###";

            }

            if (value1 == "Spend (EUR)")

            {

                //worksheet.Cells["C1"].SetStyle(style);

                pt.DataFields[1].NumberFormat = @"[>999999999]€#\,###\,###\,##0.00;[>999999]€###\,###\,##0.00;€#,###";

            }

            #endregion

            pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleDark1;



            //workbook.Worksheets[0].IsVisible = false;

            Style st = workbook.CreateStyle();
            pt.FormatAll(st);
            workbook.Save(file_location);
        }

    }

}