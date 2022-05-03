using Microsoft.Office.Interop.Excel;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using Entities;
using System.Windows.Forms;

namespace Dev_Test
{
    public class Actions
    {
        #region Unir PDF
        public void MergeXlsxFiles(string destXlsxFileName, params string[] sourceXlsxFileNames)
        {
            Application excelApp = null;
            Workbook destWorkBook = null;
            var temppathForTarget = Path.Combine(destXlsxFileName + ".xlsx");

            if (File.Exists(temppathForTarget))
                File.Delete(temppathForTarget);


            try
            {
                excelApp = new Application
                {
                    DisplayAlerts = false,
                    SheetsInNewWorkbook = 1
                };

                destWorkBook = excelApp.Workbooks.Add();
                destWorkBook.SaveAs(temppathForTarget);


                foreach (var sourceXlsxFile in sourceXlsxFileNames)
                {
                    var file = Path.Combine(Directory.GetCurrentDirectory(), sourceXlsxFile);
                    var sourceWorkBook = excelApp.Workbooks.Open(file);

                    int count = 1;
                    foreach (Excel.Worksheet ws in sourceWorkBook.Worksheets)
                    {
                        var wSheet = (Excel.Worksheet)destWorkBook.Worksheets[count];
                        ws.Copy(wSheet);
                        destWorkBook.Worksheets[destWorkBook.Worksheets.Count].Name = ws.Name + count;
                        count++;
                    }
                    sourceWorkBook.Close(XlSaveAction.xlDoNotSaveChanges);
                }

                //destWorkBook.Sheets[1].Delete();
                destWorkBook.SaveAs(Directory.GetCurrentDirectory());

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {

                if (destWorkBook != null)
                    //destWorkBook.Worksheets[destWorkBook.Worksheets.Count].Delete();
                    destWorkBook.Close(XlSaveAction.xlSaveChanges);
                if (excelApp != null)
                    excelApp.Quit();
            }
        }

        #endregion

        #region Guardar Carpeta Reciente

        #endregion

        #region Ver archivos Nuevos en la ultima carpeta

        #endregion

    }
}
