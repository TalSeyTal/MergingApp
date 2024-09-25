using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using Excel = ClosedXML.Excel;
using ClosedXML.Excel;
using System.Globalization;
using System;
using DocumentFormat.OpenXml.Packaging;

namespace MergingApp
{
    public class ExcelHelper
    {
        public DataTable ReadPolicyData(string fileName, string filePath)
        {

            DataTable dtPolicyData= new DataTable();
            dtPolicyData.Columns.Add("Number");
            dtPolicyData.Columns.Add("RenewalDate");
            dtPolicyData.Columns.Add("PaidToDate");
            dtPolicyData.Columns.Add("PolicyNumber");
            dtPolicyData.Columns.Add("FamilyName");
            dtPolicyData.Columns.Add("GivenName");
            dtPolicyData.Columns.Add("OtherGivenName");
            dtPolicyData.Columns.Add("DateOfBirth");
            dtPolicyData.Columns.Add("TFN");
            dtPolicyData.Columns.Add("TFNNotProvided");
            dtPolicyData.Columns.Add("AdmountRequested");


            var wb = new Excel.XLWorkbook(filePath + @"\" + fileName);
            var ws = wb.Worksheets.Worksheet("Sheet1");
            for (int i = 0; i < 10; i++)
            {
                DataRow drData = dtPolicyData.NewRow();

                var colA1 = ws.Cell("A"+(i+2)).Value;  //Number
                var colB1 = ws.Cell("B"+(i+2)).Value;  //Renewal Date
                var colC1 = ws.Cell("C"+(i+2)).Value;  //PaidToDate
                var colD1 = ws.Cell("D"+(i+2)).Value;  //PolicyNumber
                var colE1 = ws.Cell("E"+(i+2)).Value;  //FamilyName
                var colF1 = ws.Cell("F"+(i+2)).Value;  //GiveName
                var colG1 = ws.Cell("G"+(i+2)).Value;  //OtherGivenName
                var colH1 = ws.Cell("H"+(i+2)).Value;  //DateOfBirth
                var colI1 = ws.Cell("I"+(i+2)).Value;  //TFN
                var colJ1 = ws.Cell("j"+(i+2)).Value;   //TFNNotProvided
                var colK1 = ws.Cell("K"+(i+2)).Value;   //AdmountRequested
                drData[0] = colA1;
                drData[1] = colB1;
                drData[2] = colC1;
                drData[3] = colD1;
                drData[4] = colE1;
                drData[5] = colF1;
                drData[6] = colG1;
                drData[7] = colH1;
                drData[8] = colI1;
                drData[9] = colJ1;
                drData[10] = colK1;
                dtPolicyData.Rows.Add(drData);

            }

            int count = dtPolicyData.Rows.Count;
            return dtPolicyData;
        }

        public DataTable ReadTFNData(string fileName, string filePath)
        {

            DataTable dtTFNData = new DataTable();
            dtTFNData.Columns.Add("Policy");
            dtTFNData.Columns.Add("TFN");
            dtTFNData.Columns.Add("ROLLIN");

            var wb = new Excel.XLWorkbook(filePath + @"\" + fileName);
            var ws = wb.Worksheets.Worksheet("Sheet1");
            for (int i = 0; i < 10; i++)
            {

                DataRow drData = dtTFNData.NewRow();
                var colA1 = ws.Cell("A"+(i+2)).Value;  //Policy
                var colB1 = ws.Cell("B" + (i + 2)).Value;  //TFN
                var colC1 = ws.Cell("C" + (i + 2)).Value;  //ROLLIN

                drData[0] = colA1;
                drData[1] = colB1;
                drData[2] = colC1;
                dtTFNData.Rows.Add(drData);
            }


            int count = dtTFNData.Rows.Count;
            return dtTFNData;
        }
        public void ExporExcel(DataTable dtDataTable, string fileName)
        {

            using (XLWorkbook workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add(dtDataTable, "Sheet1");
                workSheet.Table(0).Theme = XLTableTheme.TableStyleLight20;
                workSheet.Row(1).Style.Font.Bold = true;
                workSheet.SheetView.FreezeRows(1);
                workSheet.Columns().AdjustToContents(10.0, 50.0);            
                workBook.SaveAs(@"C:\HM\"+fileName);
                    

            }

        }



       


    }
}
