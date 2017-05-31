using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;

namespace XMLToXLSService.XMLToXLSHandlers
{
    /// <summary>
    /// 
    /// </summary>
    public class MinimalMemoryLoad:IXmlToXls
    {
        private const int INITIAL_VALID_COLUMN_TO_WRITE_DATA = 1;
        private const int INITIAL_VALID_ROW_TO_WRITE_DATA = 2;

        private class WriteXLSFlags
        {
            public string LastNode { get; set; } = string.Empty;
            public int columnCount { get; set; } = INITIAL_VALID_COLUMN_TO_WRITE_DATA;
            public int rowCount { get; set; } = INITIAL_VALID_ROW_TO_WRITE_DATA;
        }


        private WriteXLSFlags Flags = new WriteXLSFlags();

        public void ExecuteXMLToXls(string filename)
        {
            XmlTextReader myTextReader = new XmlTextReader(filename);
            myTextReader.WhitespaceHandling = WhitespaceHandling.None;

            while (myTextReader.Read())
            {
                var entity = CheckNodes(myTextReader);

                if (entity != null)
                {
                    WriteExcel(entity);
                }
            }
        }

        private SimpleEntity CheckNodes(XmlTextReader xml)
        {
            if (xml.NodeType == XmlNodeType.Element && xml.IsStartElement())
            {
                Flags.LastNode = xml.LocalName;
            }

            if (xml.NodeType == XmlNodeType.Text)
            {
                var v = new SimpleEntity()
                {
                    Property = Flags.LastNode,
                    Value = xml.Value
                };

                return v;
            }

            return null;
        }


        private bool IsBeginOfFile()
        {
            return (Flags.columnCount == INITIAL_VALID_COLUMN_TO_WRITE_DATA && 
                    Flags.rowCount == INITIAL_VALID_ROW_TO_WRITE_DATA);
        }

        private void Log(string log)
        {
            //Here it goes Log4Net log code or Event Log code,...
        }

        private void WriteExcel(SimpleEntity entity)
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

            var currentDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var completeFilename = currentDir + "\\sample2.xls";

            if (IsBeginOfFile())
            {
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[1];

                if (File.Exists(completeFilename))
                    File.Delete(completeFilename);
            }
            else
            {
                if (File.Exists(completeFilename))
                {
                    xlWorkBook = xlApp.Workbooks.Open(completeFilename);
                    xlWorkSheet = xlWorkBook.Worksheets[1];
                }
                else
                {
                    xlWorkBook = xlApp.Workbooks.Add(misValue);
                    xlWorkSheet = (Worksheet)xlWorkBook.Worksheets[1];
                }
            }

            if (entity.Property.ToLower() == xlWorkSheet.Cells[1, 1].Text.ToLower())
            {
                Flags.rowCount++;
                Flags.columnCount = 1;
            }

            xlWorkSheet.Cells[1, Flags.columnCount] = entity.Property;
            xlWorkSheet.Cells[Flags.rowCount, Flags.columnCount] = entity.Value;

            Flags.columnCount++;

            xlWorkBook.SaveAs(completeFilename);

            xlWorkBook.Close(true);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Log("Excel file created , you can find the file sample2.xls");
        }


    }


}
