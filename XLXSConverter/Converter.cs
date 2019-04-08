using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;

namespace XLXSConverter
{
    public class Converter
    {

        public string temp = string.Empty;

        public void ConvertFrom()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            var to = new XLWorkbook("Bulk.xlsx");
            var bulk = to.Worksheet(1);

            for (int i = 1; i < 195; i++)
            {
                bool running = true;
                var ws = from.Worksheet(i);
                ws.Worksheet.Column(8).Cell(3);
                bulk.Column(2).Cell(1 + i).Value = ws.Worksheet.Column(8).Cell(3).Value; // sets model and make
                bulk.Column(20).Cell(1 + i).Value = ws.Worksheet.Column(8).Cell(4).Value; // sets modelperiod
                bulk.Column(16).Cell(1 + i).Value = "DK"; //Country 
                bulk.Column(17).Cell(1 + i).Value = "DK"; //Valuation country
                
                string split = ws.Worksheet.Column(8).Cell(5).Value.ToString();
                string[] splitter = split.Split(' ');
                string splitted = splitter[0];
                int category = int.Parse(splitted);

                if (category > 100 )
                {
                    bulk.Column(20).Cell(1 + i).Value = "Personal";
                }
                else
                {
                    bulk.Column(20).Cell(1 + i).Value = "Commercial";
                }

                for (int k = 4; k <= 8; k++)
                {
                    string temp = ws.Worksheet.Column(3 + k).Cell(8).Value.ToString();
                    if (temp != "")
                    {
                        bulk.Column(7).Cell(1 + k).Value = ws.Worksheet.Column(3 + k).Cell(8).Value;
                        k++;
                    }
                    else if (temp.StartsWith("Lukket"))
                    {
                        bulk.Column(7).Cell(1 + k).Value = "Van";
                    }
                    else
                    {
                        k++;
                    }

                }

                int j = 1;
                while (running)
                {

                    string temp = ws.Worksheet.Column(2).Cell(11 + j).ToString();
                    if (temp.StartsWith("Udstyr:"))
                    {
                        running = false;
                    }
                    else if (temp == "")
                    {
                        j++;
                    }
                    else
                    {
                        bulk.Worksheet.Column(6).Cell(1 + j).Value = ws.Worksheet.Column(2).Cell(11 + j).Value;
                        j++;
                    }
                }

                
            }
            
        }
    }
}
