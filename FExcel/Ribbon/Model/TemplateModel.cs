using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FExcel.FELoader.Model
{
    public class TemplateModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Mask { get; set; }
        public string FirstCellAddress { get; set; }
    }
}
