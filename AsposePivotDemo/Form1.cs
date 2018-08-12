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

            string file_location = "C:/Users/Ankita/Downloads/Europe.xlsx";



            Workbook workbook = new Workbook(file_location);



            try
            {
                workbook.Worksheets.RemoveAt("Pivot");
            }
            catch
            {

            }


            Worksheet worksheet = workbook.Worksheets.Add("Pivot");

            #region removing hard coding



            Worksheet sheet0 = workbook.Worksheets[0];

            Worksheet sheet1 = workbook.Worksheets[1];

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


            #endregion

            int iPivotIndex = worksheet.PivotTables.Add(datasource, "A1", "PivotTable");



            PivotTable pt = worksheet.PivotTables[iPivotIndex];

           

            pt.RowGrand = false;



            pt.ColumnGrand = false;



            pt.IsAutoFormat = true;

           

            pt.AddFieldToArea(PivotFieldType.Column, 0);
           // pt.ColumnHeaderCaption = "Level 1 Category";
            pt.AddFieldToArea(PivotFieldType.Row, 1);
           // pt.RowHeaderCaption = "Description";
           pt.AddFieldToArea(PivotFieldType.Data, 2);

            pt.AddFieldToArea(PivotFieldType.Data, 3);

           PivotField pf = pt.DataFields[0];
            PivotField pf1 = pt.DataFields[1];
            pt.AddFieldToArea(PivotFieldType.Column, pt.DataField);
            // PivotField pf1= pt.DataFields[1];
            // PivotFieldSubtotalType.Function = ConsolidationFunction.Sum;
            // PivotFieldSubtotalType.Function = ConsolidationFunction.Sum;
         //   pt.RowFields[0].GetSubtotals(PivotFieldSubtotalType.Sum);
           // pt.ColumnFields[0].(PivotFieldSubtotalType.Count, true);
            #region Globalisation



            Cell cell1 = sheet0.Cells["D1"];

            Cell cell2 = sheet0.Cells["C1"];

            // Style style = sheet0.Cells["C1"].GetStyle();



            //string currencySymbol = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol.ToString();

            //style.Custom = "£#,##0;[Red]-£#,##0";

            //style.Number = 5;



           string value = cell1.StringValue;

            string value1 = cell2.StringValue;
           


            if (value == "Spend USD")



            {


               workbook.Settings.LanguageCode = CountryCode.USA;

               workbook.Settings.Region = CountryCode.USA;

                //worksheet.Cells["C1"].SetStyle(style);
                pt.DataFields[1].NumberFormat = "$#,##0.00";


                // pt.DataFields[1].NumberFormat = @"[>999999999]$#\,###\,###\,##0.00;[>99999]$###\,###\,##0.00;$#,###.00";



            }



            if (value1 == "Spend (EUR)")



            {
               

               workbook.Settings.LanguageCode = CountryCode.France;

                workbook.Settings.Region = CountryCode.France;
                #region Comments
                //System.Threading.Thread.CurrentThread.CurrentCulture.ToString() ="United Kingdom";
                //worksheet.Cells["C1"].SetStyle(style);



                //pt.DataFields[0].NumberFormat= "_-[$€-2]*#.##0,00_-;-[$€-2]*#.##0,00_-;_-[$€-2]*"+"-"+"??_-;_-@_-";
                //Style style = cell2.GetStyle();
                //style.Custom= "_-[$€-2] * #.##0,00_-;-[$€-2] * #.##0,00_-;_-[$€-2] * " + " - " + "??_-;_-@_-";

                //  pt.DataFields[0].NumberFormat = @"[>999999999]£#\.###\.###\.##0,00;[>999999]£###\.###\.##0,00;£#.###,00";
                //  _ -[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_- 
                #endregion
                pt.DataFields[0].NumberFormat = "£#,##0.00";

            }



            #endregion



            pt.PivotTableStyleType = PivotTableStyleType.PivotTableStyleLight16;

         

            #region Protection level

            //workbook.Worksheets[0].IsVisible = false;

            // Restricting users to delete columns of the worksheet

            sheet1.Protection.AllowDeletingColumn = true;



            // Restricting users to delete row of the worksheet

            sheet1.Protection.AllowDeletingRow = true;



            // Restricting users to edit contents of the worksheet

            sheet1.Protection.AllowEditingContent = true;



            // Restricting users to edit objects of the worksheet

            worksheet.Protection.AllowEditingObject = true;



            // Restricting users to edit scenarios of the worksheet

            sheet1.Protection.AllowEditingScenario = true;



            // Restricting users to filter

            sheet1.Protection.AllowFiltering = true;



            // Allowing users to format cells of the worksheet

            sheet1.Protection.AllowFormattingCell = true;



            // Allowing users to format rows of the worksheet

            sheet1.Protection.AllowFormattingRow = true;



            // Allowing users to insert columns in the worksheet

            sheet1.Protection.AllowFormattingColumn = true;



            // Allowing users to insert hyper links in the worksheet

            sheet1.Protection.AllowInsertingHyperlink = true;



            // Allowing users to insert rows in the worksheet

            sheet1.Protection.AllowInsertingRow = true;



            // Allowing users to select locked cells of the worksheet

            sheet1.Protection.AllowSelectingLockedCell = true;



            // Allowing users to select unlocked cells of the worksheet

            sheet1.Protection.AllowSelectingUnlockedCell = true;



            // Allowing users to sort

            sheet1.Protection.AllowSorting = true;



            // Allowing users to use pivot tables in the worksheet

            sheet1.Protection.AllowUsingPivotTable = true;



            #endregion

            Style st = workbook.CreateStyle();

            pt.FormatAll(st);
            

            pt.RefreshData();

            pt.CalculateData();

            workbook.Save(file_location);

        }

    }

}