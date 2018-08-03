using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace csv_test_6._28._18
{
    class Read
    {
        private string File;
        public Read()
        {
            File = String.Empty;
        }
        public Read (string pFile)
        {
            File = pFile;
        }
        public string NameFile
        {
            get { return File; }
            set { File = value; }
        }        

        //////////////////////////////////////////////////////////////////////
        //                                                                  //
        //         HEADER TEXT ARRAY[]                                      //
        //                                                                  //
        //////////////////////////////////////////////////////////////////////
        public string[] TableHeader()
        {
            
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(File, false))
            {
                int j = 0;
                int rowcount = 0;
                int colcount = 0;
                

                var parts = wordDocument.MainDocumentPart.Document.Descendants<Table>().LastOrDefault();
                var row = parts.Descendants<TableRow>().FirstOrDefault();
                foreach (var cell in row.Descendants<TableCell>())
                {
                    if (!String.IsNullOrWhiteSpace(cell.InnerText))
                    {
                        colcount++;
                    }
                }
                foreach (var rownum in parts.Descendants<TableRow>())
                {                    
                    if (!String.IsNullOrWhiteSpace(rownum.Descendants<TableCell>().First().InnerText))
                    {
                        rowcount++;
                    }
                }
                string[] headertext = new string[colcount];
                var headerrow = parts.Descendants<TableRow>().First();
                foreach (var cell in headerrow.Descendants<TableCell>())
                {
                    headertext.SetValue(cell.InnerText.ToLower(), j);
                    j++;
                }
                return headertext;
            }
               
        }
        public string[,] WordDoc()
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(File, false))
            {
                int i = 0;
                int j = 0;
                int k = 5;
                int rowcount = 0;
                int colcount = 0;

                var parts = wordDocument.MainDocumentPart.Document.Descendants<Table>().LastOrDefault();
                // COUNT THE NUMBER OF COLUMNS FROM SOURCE FILE AND ASSIGN THE # TO A VARIABLE
                var row = parts.Descendants<TableRow>().FirstOrDefault();
                foreach (var cell in row.Descendants<TableCell>())
                {
                    if (!String.IsNullOrWhiteSpace(cell.InnerText))
                    {
                        colcount++;
                    }
                }
                // COUNT THE NUMBER OF ROWS FROM SOURCE FILE AND ASSIGN THE # TO A VARIABLE
                foreach (var rownum in parts.Descendants<TableRow>())
                {
                    if (!String.IsNullOrWhiteSpace(rownum.Descendants<TableCell>().First().InnerText))
                    {
                        rowcount++;                        
                    }
                }                
                string celltext = string.Empty;
                // CREATE A MULTIDIMENSIONAL ARRAY, USING THE ROW/COLUMN COUNT VARIABLES AS THE SIZE
                string[,] cellValues = new string[rowcount, colcount];
                if (parts != null)
                {
                    foreach (var node in parts.ChildElements)
                    {
                        if (node is TableRow)
                        {
                            if (parts.Descendants<TableRow>().ElementAt(0) != node)
                            {
                                if (!String.IsNullOrWhiteSpace(node.Descendants<TableCell>().FirstOrDefault().InnerText))
                                {
                                    j = 0;
                                    foreach (var cell in node.Descendants<TableCell>())
                                    {
                                        if (!String.IsNullOrWhiteSpace(cell.InnerText))
                                        {
                                            celltext = cell.InnerText.Trim();
                                            celltext.ToLower();
                                            cellValues[i, j] = celltext;
                                        }
                                        else
                                        {
                                            cellValues[i, j] = "";
                                        }
                                        j++;
                                    }
                                    i++;
                                }
                            }                 
                        }
                    }
                }
                return cellValues;
            }
        }
    }
}
