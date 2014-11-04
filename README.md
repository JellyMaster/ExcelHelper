ExcelHelper
===========

A simple project for reading and writing data to/from Excel using NPOI


For a sample export here is some quick code: 

      public ActionResult Export()
        {
            List<MyModel> model = new List<MyModel>();

            MemoryStream stream = new MemoryStream();

            model = myServiceLayer.GetExportResults();

            DataSet ds = new DataSet();
            ds.Tables.Add(model.ToDataTable<MyModel>());

            try
            {
                stream = ExcelHelper.CreateExcelSheet(ds, ExcelHelper.ExcelType.Xlsx);
            }
            catch (Exception error)
            {
                ModelState.AddModelError("", error);
            }



            if (!ModelState.IsValid)
            {
                return View();
            }
            else
            {
                string exportName = "myexcelfile.xlsx";

                return File(stream.ToArray(), "application/vnd.ms-excel", exportName);
            }



        }
