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

                if (category > 100 ) // sets category
                {
                    bulk.Column(20).Cell(1 + i).Value = "Personal";
                }
                else
                {
                    bulk.Column(20).Cell(1 + i).Value = "Commercial";
                }

                for (int k = 1; k <= 4; k++)
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

                string check = ",";
                for (int h = 1; h <= 4; h++)
                {
                    string temp = ws.Worksheet.Column(3 + h).Cell(10).Value.ToString();
                    if (temp != "")
                    {
                        bulk.Column(10).Cell(1 + h).Value = ws.Worksheet.Column(3 + h).Cell(10).Value;
                        h++;
                    }
                    else if(temp.Contains(check))
                    {
                        string[] ccm = temp.Split(',');
                        string litersToCcm = ccm[0] + ccm[1];
                        int result = int.Parse(litersToCcm);
                        int conversion = result * 1000;
                        bulk.Worksheet.Column(10).Cell(1 + h).Value = conversion.ToString();
                    }
                    else
                    {
                        h++;
                    }
                }

                for (int t = 1; t <= 4; t++)
                {
                    string temp = ws.Worksheet.Column(3 + t).Cell(11).Value.ToString();
                    if (temp != "")
                    {
                        bulk.Column(9).Cell(1 + t).Value = ws.Worksheet.Column(3 + t).Cell(11).Value;
                        t++;
                    }
                    {
                        t++;
                    }
                }

                for (int g = 1; g <= 4; g++)
                {
                    string temp = ws.Worksheet.Column(3 + g).Cell(12).Value.ToString();
                    if (temp != "")
                    {
                        bulk.Column(11).Cell(1 + g).Value = ws.Worksheet.Column(3 + g).Cell(12).Value;
                        g++;
                    }
                    {
                        g++;
                    }
                }


                
                running = true;
                int j = 1;
                int f = 1;
                while (running)
                {

                    string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
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
                        bulk.Worksheet.Column(6).Cell(1 + f).Value = ws.Worksheet.Column(4).Cell(12 + j).Value;
                        bulk.Worksheet.Column(22).Cell(1 + f).Value = ws.Worksheet.Column(4).Cell(12 + j).Value;
                        bulk.Worksheet.Column(21).Cell(1 + f).Value = ws.Worksheet.Column(4).Cell(13 + j).Value;
                        j++;
                        f++;
                    }
                }

                running = true;
                j = 1;
                f = 1;
                while (running)
                {

                    string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
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
                        bulk.Worksheet.Column(6).Cell(1 + f).Value = ws.Worksheet.Column(5).Cell(12 + j).Value;
                        bulk.Worksheet.Column(22).Cell(1 + f).Value = ws.Worksheet.Column(5).Cell(12 + j).Value;
                        bulk.Worksheet.Column(21).Cell(1 + f).Value = ws.Worksheet.Column(5).Cell(13 + j).Value;
                        j++;
                        f++;
                    }
                }

                running = true;
                j = 1;
                while (running)
                {

                    string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
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
                        bulk.Worksheet.Column(6).Cell(1 + f).Value = ws.Worksheet.Column(6).Cell(12 + j).Value;
                        bulk.Worksheet.Column(22).Cell(1 + f).Value = ws.Worksheet.Column(6).Cell(12 + j).Value;
                        bulk.Worksheet.Column(21).Cell(1 + f).Value = ws.Worksheet.Column(6).Cell(13 + j).Value;
                        j++;
                        f++;
                    }
                }

                running = true;
                j = 1;
                while (running)
                {

                    string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
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
                        bulk.Worksheet.Column(6).Cell(1 + f).Value = ws.Worksheet.Column(7).Cell(12 + j).Value;
                        bulk.Worksheet.Column(22).Cell(1 + f).Value = ws.Worksheet.Column(7).Cell(12 + j).Value;
                        bulk.Worksheet.Column(21).Cell(1 + f).Value = ws.Worksheet.Column(7).Cell(13 + j).Value;
                        j++;
                        f++;
                    }
                }

                running = true;
                j = 1;
                while (running)
                {

                    string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
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
                        bulk.Worksheet.Column(6).Cell(1 + f).Value = ws.Worksheet.Column(8).Cell(12 + j).Value;
                        bulk.Worksheet.Column(22).Cell(1 + f).Value = ws.Worksheet.Column(8).Cell(12 + j).Value;
                        bulk.Worksheet.Column(21).Cell(1 + f).Value = ws.Worksheet.Column(8).Cell(13 + j).Value;
                        j++;
                        f++;
                    }
                }

                if ()
                {

                }


            }
            
        }
    }
}
