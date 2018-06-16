using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace _20180616_COMExcelObjLibrary
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application xlApp = Globals.ThisAddIn.Application;
            Excel.Workbook xlWb = xlApp.ActiveWorkbook;
            Excel.Worksheet xlSh = xlApp.Worksheets["TEST"];
            Excel.ListObject xlListObj = xlSh.ListObjects["TEST"];

            Excel.Sort xlListObjSort = xlListObj.Sort;
            Excel.SortFields xlListObjSortFields = xlListObjSort.SortFields;

            xlListObjSortFields.Clear();
        }
    }
}
