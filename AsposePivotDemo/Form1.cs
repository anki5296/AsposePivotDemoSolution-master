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

            MemoryStream mems = new MemoryStream();


            

            string file_location = @"C:/Users/Ankita/Downloads/subtotals.xlsx";

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


            string datasource = sheetname + "!A5:" + columnName + row_last.ToString();

            #region JSON


            #endregion

            int skiprows = 5;
            int iPivotIndex = worksheet.PivotTables.Add(datasource, "A" + skiprows.ToString(), "PivotTable");
            PivotTable pt = worksheet.PivotTables[iPivotIndex];
            pt.RowGrand = true;
           
            pt.ColumnGrand = true;

            pt.ShowDrill = true;
           
            pt.IsAutoFormat = true;
            //int legendctr = 2;

            //while (legendctr < skiprows)
            //{
            //    if (true)
            //    {
            //        Cell report_title_cell = sheet0.Cells["A" + legendctr.ToString()];
            //        string report = report_title_cell.Value.ToString();
            //        Cell pivot_report_title_cell = sheet1.Cells["A" + legendctr.ToString()];
            //        legendctr++;
            //        pivot_report_title_cell.PutValue(report);
            //    }
            //    if (true)
            //    {
            //        Cell execution_date_cell = sheet0.Cells["A" + legendctr.ToString()];
            //        string execution = execution_date_cell.Value.ToString();
            //        Cell pivot_report_title_cell = sheet1.Cells["A" + legendctr.ToString()];
            //        legendctr++;

            //        pivot_report_title_cell.PutValue(execution);
            //    }
            //    if (true)
            //    {
            //        Cell filter_cell = sheet0.Cells["A" + legendctr.ToString()];
            //        string filter1 = filter_cell.Value.ToString();
            //        Cell pivot_report_title_cell = sheet1.Cells["A" + legendctr.ToString()];
            //        legendctr++;

            //        pivot_report_title_cell.PutValue(filter1);
            //    }



            //}

            ////  Cell cell_title = sheet0.Cells["A" + 2.ToString()]; 

            //      //  Cell cell_title = sheet0.Cells["A2"];



            ////  Cell cell_filter = sheet0.Cells["A3"];
            //  //string title;


            //  title = cell_title.Value.ToString();

            //  string filter;
            //  filter = cell_filter.Value.ToString();

            // Cell pivot_cell_title = sheet1.Cells["A1"];
            //  Cell pivot_cell_filter = sheet1.Cells["A2"];
            //  pivot_cell_title.PutValue(title);
            //  pivot_cell_filter.PutValue(filter);
            //  // int counter = 0;
            /**using (StreamReader r = new StreamReader(@"C:\Users\Ankita\Downloads\AsposePivotDemoSolution-master-master\AsposePivotDemo\Report1.json"))
            {
                string json = r.ReadToEnd();

                dynamic array = JsonConvert.DeserializeObject(json);
                try
                {
                    int i = 0;
                    if (array["LstReportObjectOnColumn"] != null)
                    {

                        foreach (var item in array["LstReportObjectOnColumn"])
                        {
                            pt.AddFieldToArea(PivotFieldType.Column, item["ReportObjectName"].ToString());
                            //Console.WriteLine(item["ReportObjectName"]);
                        }

                    }

                    if (array["LstReportObjectOnRow"] != null)
                    {

                        foreach (var item in array["LstReportObjectOnRow"])
                        {
                            pt.AddFieldToArea(PivotFieldType.Row, item["ReportObjectName"].ToString());
                            //workbook.Worksheets[1].PivotTables[0].RowFields[i++].IsAutoSubtotals = false;

                        }

                    }

                    if (array["LstReportObjectOnValue"] != null)
                    {
                        foreach (var item in array["LstReportObjectOnValue"])
                        {
                            pt.AddFieldToArea(PivotFieldType.Data, item["ReportObjectName"].ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    ex.ToString();
                }
            }**/

         pt.AddFieldToArea(PivotFieldType.Row, 0);
          pt.AddFieldToArea(PivotFieldType.Column, 1);
            //pt.AddFieldToArea(PivotFieldType.Column, 1);
          pt.AddFieldToArea(PivotFieldType.Data, 2);
            //pt.DataFields[0].Number = 10;
            pt.ColumnFields[0].Number = 14;

            //CellArea ca = CellArea.CreateCellArea("A3", "D163");


            //DataSorter sorter = workbook.DataSorter;
            //int idx = CellsHelper.ColumnNameToIndex("A");
            //sorter.AddKey(idx, Aspose.Cells.SortOrder.Ascending);
            //sorter.Sort(sheet0.Cells, ca);






            // pt.ColumnHeaderCaption = "Level 1 Category";

            // pt.RowHeaderCaption = "Description";
            //pt.AddFieldToArea(PivotFieldType.Data, 2);

            //pt.AddFieldToArea(PivotFieldType.Data, 3);

            //PivotField pf = pt.DataFields[0];
            // PivotField pf1 = pt.DataFields[1];
            //pt.AddFieldToArea(PivotFieldType.Column, pt.DataFields
            // PivotField pf1= pt.DataFields[1];
            // PivotFieldSubtotalType.Function = ConsolidationFunction.Sum;
            // PivotFieldSubtotalType.Function = ConsolidationFunction.Sum;
            //   pt.RowFields[0].GetSubtotals(PivotFieldSubtotalType.Sum);
            // pt.ColumnFields[0].(PivotFieldSubtotalType.Count, true);
            #region Globalisation



            //Cell cell1 = sheet0.Cells["D1"];

            //Cell cell2 = sheet0.Cells["C1"];

            // Style style = sheet0.Cells["C1"].GetStyle();



            //string currencySymbol = System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.CurrencySymbol.ToString();

            //style.Custom = "£#,##0;[Red]-£#,##0";

            //style.Number = 5;



            //string value = cell1.StringValue;

            // string value1 = cell2.StringValue;



            // if (value == "Spend USD")



            // {


            //    workbook.Settings.LanguageCode = CountryCode.USA;

            //    workbook.Settings.Region = CountryCode.USA;

            //     //worksheet.Cells["C1"].SetStyle(style);
            //     pt.DataFields[1].NumberFormat = "$#,##0.00";


            //     // pt.DataFields[1].NumberFormat = @"[>999999999]$#\,###\,###\,##0.00;[>99999]$###\,###\,##0.00;$#,###.00";



            // }



            //if (value1 == "Spend (EUR)")



            //{


            //   workbook.Settings.LanguageCode = CountryCode.France;

            //    workbook.Settings.Region = CountryCode.France;
            //    #region Comments
            //    //System.Threading.Thread.CurrentThread.CurrentCulture.ToString() ="United Kingdom";
            //    //worksheet.Cells["C1"].SetStyle(style);



            //    //pt.DataFields[0].NumberFormat= "_-[$€-2]*#.##0,00_-;-[$€-2]*#.##0,00_-;_-[$€-2]*"+"-"+"??_-;_-@_-";
            //    //Style style = cell2.GetStyle();
            //    //style.Custom= "_-[$€-2] * #.##0,00_-;-[$€-2] * #.##0,00_-;_-[$€-2] * " + " - " + "??_-;_-@_-";

            //    //  pt.DataFields[0].NumberFormat = @"[>999999999]£#\.###\.###\.##0,00;[>999999]£###\.###\.##0,00;£#.###,00";
            //    //  _ -[$€-2] * #,##0.00_-;-[$€-2] * #,##0.00_-;_-[$€-2] * "-"??_-;_-@_- 
            //    #endregion
            //    pt.DataFields[0].NumberFormat = "£#,##0.00";

            //}



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
            //chganss

            sheet1.Protection.AllowSorting = true;



            // Allowing users to use pivot tables in the worksheet

            sheet1.Protection.AllowUsingPivotTable = true;



            #endregion

            // Style st = workbook.CreateStyle();

            //     pt.FormatAll(st);

            // pt.AutoFormatType =PivotTableAutoFormatType.Table10;
           // workbook.Worksheets[1].PivotTables[0].RowFields[0].IsAutoSubtotals = false;

            pt.ColumnFields[0].IsAutoSubtotals = false;
            

            //pt.RefreshData();

            pt.CalculateData();

             workbook.Save(file_location);
         


            //workbook.Save(mems, SaveFormat.Xlsx);
            
            //Workbook wb2 = new Workbook(mems);
            //Worksheet ws2 = wb2.Worksheets.Add("Pivot123");

            //wb2.Save(mems, SaveFormat.Xlsx);

            ////MemoryStream newms = new MemoryStream();
            //// mems.Seek(0, SeekOrigin.Begin);
            // mems.Position = 0;
            //using (FileStream fs = new FileStream(@"C:/Users/Ankita/Downloads/plswork1.xlsx", FileMode.OpenOrCreate))
            //{
            //   mems.CopyTo(fs);
            //   fs.Flush();
            //}
        }

    }

}
