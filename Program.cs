using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace MergingApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelHelper excelHelper = new ExcelHelper();
            DataTable dtPolicyData= excelHelper.ReadPolicyData("PolicyData.xlsx",@"c:\HM");
            DataTable dtTFNData = excelHelper.ReadTFNData("TFN.xlsx", @"c:\HM");

            for (int i = 0; i < dtPolicyData.Rows.Count; i++)
            {
                DataRow drPolicyData = dtPolicyData.Rows[i];
                string policyNumber = drPolicyData[3].ToString();
                for (int j = 0; j< dtTFNData.Rows.Count; j++)
                {
                    DataRow drTFNData = dtTFNData.Rows[j];
                    string tfnPolicyNum= drTFNData[0].ToString();
                    string tfnNumber = drTFNData[1].ToString();
                    string amount = drTFNData[2].ToString();
                    if (policyNumber.Equals(tfnPolicyNum))
                    {
                        drPolicyData[8] = tfnNumber;
                        drPolicyData[10] = amount;
                        break;
                    }
                }
            }
            DataRow drPolicyData1 = dtPolicyData.Rows[0];
            string tfnUmber1 = drPolicyData1[8].ToString();
            string amount1 = drPolicyData1[10].ToString();

            excelHelper.ExporExcel(dtPolicyData, "FinalOutput.xlsx");

        }
    }
}
