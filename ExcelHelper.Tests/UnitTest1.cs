using System;
using System.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Diagnostics;

namespace ExcelHelper.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void SingleSheet()
        {

            MemoryStream stream = new MemoryStream(File.ReadAllBytes(@"Book1-one sheet.xlsx"));



            DataSet dataset = ExcelHelper.CreateDataSetFromExcel(stream, ExcelHelper.ExcelType.Xlsx.ToString());


            Debug.WriteLine(string.Format("Found Number of columns: {0}", dataset.Tables[0].Columns.Count));
            Debug.WriteLine(string.Format("Found Number of rows: {0}", dataset.Tables[0].Rows.Count));




            Assert.IsTrue(dataset.Tables.Count == 1, "more or less than one table found");



        }


        [TestMethod]
        public void MultiSheet()
        {

            MemoryStream stream = new MemoryStream(File.ReadAllBytes(@"Book1 - multi sheet.xlsx"));



            DataSet dataset = ExcelHelper.CreateDataSetFromExcel(stream,ExcelHelper.ExcelType.Xlsx.ToString() );


            Debug.WriteLine(string.Format("Found Number of columns: {0} in table 1", dataset.Tables[0].Columns.Count));
            Debug.WriteLine(string.Format("Found Number of rows: {0} int table 1", dataset.Tables[0].Rows.Count));

            Debug.WriteLine(string.Format("Found Number of columns: {0} in table 2", dataset.Tables[1].Columns.Count));
            Debug.WriteLine(string.Format("Found Number of rows: {0} in table 2", dataset.Tables[1].Rows.Count));




            Assert.IsTrue(dataset.Tables.Count == 2, "more or less than one table found");



        }




        [TestMethod]
        public void CreateSheetFromData()
        {
            MemoryStream stream = new MemoryStream(File.ReadAllBytes(@"Book1-one sheet.xlsx"));



            DataSet dataset = ExcelHelper.CreateDataSetFromExcel(stream, "xlsx");

            MemoryStream stream2 = new MemoryStream();

            stream2 = ExcelHelper.CreateExcelSheet(dataset);




            File.WriteAllBytes(@"savedfile.xlsx", stream2.GetBuffer());

            FileInfo file = new FileInfo(@"savedfile.xlsx");

            Debug.WriteLine(file.FullName);

            Assert.IsTrue(File.Exists(@"savedfile.xlsx"));
        }


        

    }
}
