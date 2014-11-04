ExcelHelper
===========

A simple project for reading and writing data to/from Excel using NPOI


For a sample export here is some quick code: 

      public ActionResult Export()
        {
            //The collection we want to convert into an excel file 
            List<MyModel> model = new List<MyModel>();

            MemoryStream stream = new MemoryStream();
            //Some form of service layer to get the results collection. 
            model = myServiceLayer.GetExportResults();

            DataSet ds = new DataSet();
            
            //convert the list to a datatable (Note: must use simple properties. Complex properties are not supported currently)
            ds.Tables.Add(model.ToDataTable<MyModel>());

            try
            {
            //Create the excel sheet and write it to the memory stream. 
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
                //return the file. 

                return File(stream.ToArray(), "application/vnd.ms-excel", exportName);
            }



        }
