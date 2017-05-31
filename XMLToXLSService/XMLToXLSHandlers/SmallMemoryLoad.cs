using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XMLToXLSService.XMLToXLSHandlers
{
    /// <summary>
    /// This class is faster and works by gathering a small bunch of entities in memory before running the xls conversion
    /// </summary>
    public class SmallMemoryLoad: IXmlToXls
    {
        public void ExecuteXMLToXls(string filename)
        {
            XmlTextReader myTextReader = new XmlTextReader(filename);
            myTextReader.WhitespaceHandling = WhitespaceHandling.None;
            while (myTextReader.Read())
            {
                CheckXmlNodes(myTextReader);
            }

            ExcellInterop(Entities);
        }

        private void CheckXmlNodes(XmlTextReader xml)
        {
            if (xml.NodeType == XmlNodeType.Element && xml.IsStartElement())
            {
                LastNode = xml.LocalName;
            }

            if (xml.NodeType == XmlNodeType.Text)
            {
                var v = new SimpleEntity()
                {
                    Property = LastNode,
                    Value = xml.Value
                };

                Entities.Add(v);
            }
        }

        private string LastNode = string.Empty;
        private List<SimpleEntity> Entities = new List<SimpleEntity>();

        private int Rows()
        {
            return Entities.Count(x => x.Property == Entities[0].Property);
        }

        private int Columns()
        {
            if (Rows() > 0)
                return Entities.Count() / Rows();
            return 0;
        }

        private void Log(string log)
        {
            //Here it goes Log4Net log code or Event Log code,...
        }

        private void ExcellInterop(List<SimpleEntity> entities)
        {
            Application xlApp = new Application();

            if (xlApp == null)
            {
                Log("Excel is not properly installed!!");
                return;
            }

            xlApp.DisplayAlerts = false;

            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[1];

            var columns = Columns();
            var rows = Rows();


            for (int i = 0; i < columns; i++)
            {
                xlWorkSheet.Cells[1, i + 1] = entities[i].Property;
            }


            int currentRow = 1;
            int currentColumn = 0;
            for (int i = 0; i < entities.Count; i++)
            {
                currentColumn++;
                if (i % columns == 0)
                {
                    currentRow++;
                    currentColumn = 1;
                }

                xlWorkSheet.Cells[currentRow, currentColumn] = entities[i].Value;
            }

            var currentDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            xlWorkBook.SaveAs(currentDir + "\\sample.xls");
            xlWorkBook.Close(true);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Log("Excel file created , you can find the file sample.xls");
        }

    }

}
