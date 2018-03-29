using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnStart_Click(object sender, RibbonControlEventArgs e)
        {
            //open excel worksheet
            Excel.Worksheet activeWorkSheet = Globals.ThisAddIn.Application.ActiveSheet;
            //get cell
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
          //  Type type = activeWorkSheet.Cells.Value.GetType();
            if(actCell.Value != null )
            {
                Type type = actCell.Value.GetType();
                MessageBox.Show("The type data of the cell is " + type);
            }
        }
    }
}
