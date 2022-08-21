using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FExcel.FELoader.Model
{
    public class ParamModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public bool IsMFSO { get; set; }
        public bool IsSelected { get; set; }
        public Dictionary<string, string> Formula { get; set; }
        public int RowID { get; set; }
    }
}
