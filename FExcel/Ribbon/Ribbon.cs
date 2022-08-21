using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
namespace FExcel
{
    [ComVisible(true)]
    public class RibbonManager : ExcelRibbon
    {
        public void OnShowCTP(IRibbonControl control)
        {
            CTPManager.ShowCTP();
        }

        public void OnDeleteCTP(IRibbonControl control)
        {
            CTPManager.DeleteCTP();
        }
    }
}
