using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FExcel.FELoader.Model
{
    public class FileElement
    {
        public int Id { get; set; }
        public string FilePath { get; set; }
        public string BookName { get; set; }
        public string BookStructure { get; set; }
        public string SheetName { get; set; }
        public string TemplateName { get; set; }
        public int ShiftYear { get; set; }
        public string Group { get; set; }
        public string OG { get; set; }
        public string Mest { get; set; }
        public double MFSOKoeff { get; set; }
        public bool IsSelected { get; set; }

        /// <summary>
        /// Получить значение текущей категории для группировки
        /// </summary>
        /// <param name="element">текущий элемент FileElement</param>
        /// <returns></returns>
        public static string GetGroupBy(FileElement element)
        {
            var res = element.OG;
            var curGroup = Properties.Settings.Default.GroupBy;
            var curFileElementType = Enum.Parse(typeof(FileElementType), curGroup);

            switch (curFileElementType)
            {
                case FileElementType.FilePath:
                    res = element.FilePath;
                    break;
                case FileElementType.BookName:
                    res = element.BookName;
                    break;
                case FileElementType.BookStructure:
                    res = element.BookStructure;
                    break;
                case FileElementType.SheetName:
                    res = element.SheetName;
                    break;
                case FileElementType.TemplateName:
                    res = element.TemplateName;
                    break;
                case FileElementType.Group:
                    res = element.Group;
                    break;
                case FileElementType.OG:
                    res = element.OG;
                    break;
                case FileElementType.Mest:
                    res = element.Mest;
                    break;
                case FileElementType.MFSOKoeff:
                    res = element.MFSOKoeff.ToString();
                    break;
                default:
                    break;
            }

            return res;
        }
    }

    public enum FileElementType
    {
        FilePath, BookName, BookStructure, SheetName, TemplateName, Group, OG, Mest, MFSOKoeff
    }
}
