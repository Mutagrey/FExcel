using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FExcel.FELoader.Model
{
    public class Category
    {
        public string CategoryName { get; set; }
        public int CategoryCount { get; set; }
        public string CategoryDescription { get; set; }

        public Category(string categoryName, int categoryCount = 0, string categoryDescription = "")
        {
            CategoryName = categoryName;
            CategoryCount = categoryCount;
            CategoryDescription = categoryDescription;
        }

    }
}
