using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RevitPluginsApp.Plugin.ClashManagement
{
    public class HtmlReportParser
    {
        public IHtmlDocument GetHtmlDocument(string fileName)
        {
            string htmlFile = File.ReadAllText(fileName, Encoding.UTF8);

            return new HtmlParser().ParseDocument(htmlFile);
        }

        public string GetReportName(IHtmlDocument htmlDoc)
        {
            return htmlDoc.QuerySelector(".testName")?.InnerHtml ?? "Неизвестный отчет";
        }

        public IElement GetReportMainTable(IHtmlDocument htmlDoc)
        {
            return htmlDoc.QuerySelector(".mainTable");
        }

        public IHtmlCollection<IElement> GetMainTableRows(IElement mainTable)
        {
            var mainTableSection = mainTable.Children.FirstOrDefault();

            return mainTableSection.Children;
        }

        public IHtmlCollection<IElement> GetMainTableHeaderColumns(IHtmlCollection<IElement> mainTableRows)
        {
            var headerRows = mainTableRows.Where(i => i.ClassName == "headerRow").ToList();

            var headerRow = headerRows[1];

            return headerRow.Children;
        }
        public List<IElement> GetGeneralHeaderColumns(IHtmlCollection<IElement> headerColumns)
        {
            return headerColumns.Where(c => c.ClassName == "generalHeader").ToList();
        }

        public List<IElement> GetItemHeaderColumns(IHtmlCollection<IElement> headerColumns)
        {
            return headerColumns.Where(c => c.ClassName == "item1Header").ToList();
        }

        public List<IElement> GetTableContentRows(IHtmlCollection<IElement> mainTableRows)
        {
            return mainTableRows.Where(i => i.ClassName == "contentRow").ToList();
        }

        public int GetColumnIndex(List<IElement> columns, string columnName)
        {
            return columns.FindIndex(c => c.InnerHtml == columnName);
        }
    }
}
