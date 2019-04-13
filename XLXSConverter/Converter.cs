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
        public List<string> MakeModel = new List<string>();
        public List<string> Category = new List<string>();
        public List<string> FirstReg = new List<string>();
        public List<string> Body = new List<string>();
        public List<string> Engine = new List<string>();
        public List<string> EngineLiters = new List<string>();
        public List<string> Fuel = new List<string>();
        public List<string> Odometer = new List<string>();
        public List<string> ModelPeriode = new List<string>();
        public List<string> NewPrice = new List<string>();
        public List<string> ListPrice = new List<string>();


        public string country = "DK";
        public string CatPer = "Personal";
        public string CatCom = "Commercial";

        public void StoreMakeModel()
        {
            // sets model and make
            var from = new XLWorkbook("Niveaulister1.xlsx");

            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);

                MakeModel.Add(ws.Worksheet.Column(8).Cell(2).Value.ToString());

            }
        }

        public void StoreModelPeriode()
        {
            // sets modelperiod - bulk.Column(20).Cell(1 + i).Value
            var from = new XLWorkbook("Niveaulister1.xlsx");
            for (int i = 2; i <= 195; i++)
            {

                var ws = from.Worksheet(i);
                ModelPeriode.Add(ws.Worksheet.Column(8).Cell(4).Value.ToString());
            }
        }

        public void StoreEngineLiters()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);
                string check = ",";
                for (int h = 1; h <= 4; h++)
                {
                    string temp = ws.Worksheet.Column(3 + h).Cell(10).Value.ToString();
                    if (temp != "")
                    {
                        EngineLiters.Add(ws.Worksheet.Column(3 + h).Cell(10).Value.ToString());
                        h++;
                    }
                    else if (temp.Contains(check))
                    {
                        string[] ccm = temp.Split(',');
                        string litersToCcm = ccm[0] + ccm[1];
                        int result = int.Parse(litersToCcm);
                        int conversion = result * 1000;
                        EngineLiters.Add(conversion.ToString());
                    }
                    else
                    {
                        h++;
                    }
                }
            }
        }

        public void StoreBody()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);
                for (int k = 1; k <= 4; k++)
                {
                    string temp = ws.Worksheet.Column(3 + k).Cell(8).Value.ToString();
                    if (temp != "")
                    {
                        Body.Add(ws.Worksheet.Column(3 + k).Cell(8).Value.ToString());
                        k++;
                    }
                    else if (temp.StartsWith("Lukket"))
                    {
                        Body.Add("Van");
                    }
                    else
                    {
                        k++;
                    }
                }
            }
        }

        public void StorFuelType()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);
                for (int g = 1; g <= 4; g++)
                {
                    string temp = ws.Worksheet.Column(3 + g).Cell(12).Value.ToString();
                    if (temp != "")
                    {
                        Fuel.Add(ws.Worksheet.Column(3 + g).Cell(12).Value.ToString());
                        g++;
                    }
                    {
                        g++;
                    }
                }

            }
        }
        public void StoreFirstReg()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            int j = 1;
            int y = 1;
            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);

                if (y > 230)
                {
                    string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
                    if (temp.StartsWith("Udstyr:"))
                    {
                        y = 230;
                    }
                    else if (temp == "")
                    {
                        j++;
                        y++;
                    }
                    else
                    {
                        FirstReg.Add(ws.Worksheet.Column(2).Cell(12 + j).Value.ToString());
                        j++;
                        y++;
                    }
                }
            }
        }

        public void StoreEngine()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);
                for (int t = 1; t <= 5; t++)
                { //Engine
                    string temp = ws.Worksheet.Column(3 + t).Cell(11).Value.ToString();
                    if (temp != "" && temp != "0")
                    {
                        Engine.Add(ws.Worksheet.Column(3 + t).Cell(11).Value.ToString());
                        t++;
                    }
                    else
                    {
                        t++;
                    }
                }

            }
        }

        public void StoreCategory()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");


            for (int i = 2; i < 195; i++)
            {
                var ws = from.Worksheet(i);
                string split = ws.Worksheet.Column(8).Cell(5).Value.ToString();
                string[] splitter = split.Split('%', '/');
                string splitted = splitter[0].ToString();
                //string final = splitted.Trim('%');
                int category = int.Parse(splitted);

                if (category == 100) // sets category
                {
                    Category.Add(CatPer);
                }
                else if (category == 50)
                {
                    Category.Add(CatCom);
                }
                else Category.Add("Partially commercial");
            }

        }

        public void StoreNewPrice()
        {
            var from = new XLWorkbook("Niveaulister1.xlsx");
            for (int i = 2; i <= 195; i++)
            {
                var ws = from.Worksheet(i);
                int j = 1;
                int f = 1;
                for (int y = 0; y < 230; y++)

                {

                    j = 1;
                    f = 1;
                    for (int k = 0; k < 230; k++)
                    {

                        string temp = ws.Worksheet.Column(2).Cell(12 + j).ToString();
                        if (temp.StartsWith("Udstyr:"))
                        {
                            k = 230;
                        }
                        else if (temp == "")
                        {
                            j++;
                            k++;
                        }
                        else
                        {
                            ListPrice.Add(ws.Worksheet.Column(5).Cell(12 + j).Value.ToString());
                            NewPrice.Add(ws.Worksheet.Column(5).Cell(12 + j + f).Value.ToString());
                            j++;
                            k++;

                        }
                    }
                }
            }
        }

        public void WriteTo()
        {
            var to = new XLWorkbook("Bulk.xlsx");
            var bulk = to.Worksheet(1);


            foreach (string item in Engine)
            {
                Console.WriteLine(item);
            }

            //bulk.Column(16).Cell(1 + i).Value = country; //Country 
            //bulk.Column(17).Cell(1 + i).Value = country; //Valuation country
        }
    }

}