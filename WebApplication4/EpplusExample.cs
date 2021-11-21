using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace EPPlusSamples.LoadingData
{
    [EpplusTable(TableStyle = TableStyles.None, PrintHeaders = true, AutofitColumns = true)]
    internal class Actor
    {
        [EpplusIgnore]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 3)]
        public string LastName { get; set; }
        [EpplusTableColumn(Order = 1, Header = "First name")]
        public string FirstName { get; set; }
        [EpplusTableColumn(Order = 2)]
        public string MiddleName { get; set; }

        [EpplusTableColumn(Order = 0, NumberFormat = "yyyy-MM-dd")]
        public DateTime Birthdate { get; set; }
    }


    public static class LoadingDataFromCollectionWithAttributes
    {
        public static void Run()
        {
            FileUtil.OutputDir = new System.IO.DirectoryInfo("C:\\test");
            // sample data
            var actors = new List<Actor>
            {
                new Actor{ FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
                new Actor{ FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
                new Actor{ FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
            };

         

            using (var package = new ExcelPackage(FileUtil.GetCleanFileInfo("04-LoadFromCollectionAttributes.xlsx")))
            {
                List<TableStyles> styles = new List<TableStyles>();
                // using the Actor class above
                styles.Add(TableStyles.Light1);
                styles.Add(TableStyles.Light2);
                styles.Add(TableStyles.Light3);
                styles.Add(TableStyles.Light4);
                styles.Add(TableStyles.Light5);
                styles.Add(TableStyles.Light6);
                styles.Add(TableStyles.Light7);
                styles.Add(TableStyles.Light8);
                styles.Add(TableStyles.Light9);
                styles.Add(TableStyles.Light10);
                styles.Add(TableStyles.Light11);
                styles.Add(TableStyles.Light12);
                styles.Add(TableStyles.Light13);
                styles.Add(TableStyles.Light14);
                styles.Add(TableStyles.Light15);
                styles.Add(TableStyles.Light16);
                styles.Add(TableStyles.Light17);
                styles.Add(TableStyles.Light18);
                styles.Add(TableStyles.Light19);
                styles.Add(TableStyles.Light20);
                styles.Add(TableStyles.Light21);

                //for (int i = 0; i < 21; i++)
                //{
                var sheet = package.Workbook.Worksheets.Add("Actors");
                ExcelRangeBase rangeBase = sheet.Cells["B2"].LoadFromCollection(actors);
                ExcelRange Rng = sheet.Cells[rangeBase.Address];


                ExcelTableCollection tblcollection = sheet.Tables;
                ExcelTable table = tblcollection.GetFromRange(Rng);
                //table.TableStyle = styles[17];
                table.ShowFilter = false;
                table.ShowRowStripes = false;
                table.HeaderRowStyle.Fill.PatternType = ExcelFillStyle.Solid;
                table.HeaderRowStyle.Fill.BackgroundColor.SetColor(Color.LightGray);
                table.HeaderRowStyle.Font.Bold = true;
                table.HeaderRowStyle.Border.Top.Style = ExcelBorderStyle.Thin;
                table.HeaderRowStyle.Border.Top.Color.SetColor(Color.Black);
                table.HeaderRowStyle.Border.Left.Style = ExcelBorderStyle.Thin;
                table.HeaderRowStyle.Border.Left.Color.SetColor(Color.Black);
                table.HeaderRowStyle.Border.Right.Style = ExcelBorderStyle.Thin;
                table.HeaderRowStyle.Border.Right.Color.SetColor(Color.Black);
                table.HeaderRowStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
                table.HeaderRowStyle.Border.Bottom.Color.SetColor(Color.Black);

                table.DataStyle.Border.Top.Style = ExcelBorderStyle.Thin;
                table.DataStyle.Border.Top.Color.SetColor(Color.Black);
                table.DataStyle.Border.Left.Style = ExcelBorderStyle.Thin;
                table.DataStyle.Border.Left.Color.SetColor(Color.Black);
                table.DataStyle.Border.Right.Style = ExcelBorderStyle.Thin;
                table.DataStyle.Border.Right.Color.SetColor(Color.Black);
                table.DataStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
                table.DataStyle.Border.Bottom.Color.SetColor(Color.Black);
                //Rng.Style.Border.BorderAround(ExcelBorderStyle.Thin);

                ExcelRange title = sheet.Cells[1, 2, 1, table.Columns.Count + 1];
                title.Merge = true;
                title.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                title.Style.Font.Bold = true;
                title.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                title.Value = "Test Tile";


            // }

               
                package.Save();
            }
        }
    }
}