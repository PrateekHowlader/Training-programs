namespace bom_of_teckla_meterials
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using ClosedXML.Excel;
    using iTextSharp.text;
    using iTextSharp.text.pdf;
    using DocumentFormat.OpenXml.Spreadsheet;
    using DocumentFormat.OpenXml.Drawing.Charts;
    using static System.Runtime.InteropServices.JavaScript.JSType;
    using System.Text.RegularExpressions;
    using System.Diagnostics.Metrics;

    public class MaterialItem
    {
        public string Profile { get; set; }
        public string Grade { get; set; }
        public string Qty { get; set; }
        public string Length { get; set; }
        public string Area_in2   { get; set; }
        public string Weight { get; set; }
        
        public MaterialItem(string profile , string grade , string qty ,string length , string area , string weight) 
        {
            Profile = profile;
            Grade = grade;
            Qty = qty;
            Length = length;
            Area_in2 = area;
            Weight = weight;
        }
        public void Display()
        {
            Console.WriteLine(Profile+"\t\t"+ Grade + "\t\t" + Qty + "\t\t" + Length + "\t\t" + Area_in2 + "\t\t" + Weight);
        }
    }

    class Program
    {
        static void Main()
        {
            Console.Write(" enter the exel address : ");
            string csvPath =Console.ReadLine() ;
            Console.Write(" enter the new address : ");
            string new_adderss = Console.ReadLine();
            List<MaterialItem> materialItems = ReadMaterialListFromCsv(csvPath) ;
            
            Console.WriteLine(materialItems.Capacity);
            CreatePdfFromList(materialItems, new_adderss);
            //CreateExcelFile(new_adderss, materialItems);
            // List<MaterialItem> materials = ReadMaterialListFromCsv(csvPath);
            // GenerateBomReport(materials, csvPath);
            Console.WriteLine("BOM Report generated successfully.");
        }

        static List<MaterialItem> ReadMaterialListFromCsv(string csvPath)
        {
            
            List<MaterialItem> materials_list = new List<MaterialItem>();
            using (var workbook = new XLWorkbook(csvPath))
            {
                var worksheet = workbook.Worksheet(1); // Read from the first worksheet
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip the header row

                foreach (var row in rows)
                {
                    string profile = row.Cell(1).GetString(); // Read cell 1
                    string grade = row.Cell(2).GetString(); // Read cell 2
                    string qty = row.Cell(3).GetString(); // Read cell 3
                    string length = row.Cell(4).GetString();
                    string area = row.Cell(5).GetString();
                    string weight = row.Cell(6).GetString();


                    materials_list.Add(new MaterialItem(profile, grade, qty, length, area, weight));
                }
                foreach (var material in materials_list)
                {
                    material.Display();
                }
            }

            return materials_list;
        }
    
        public static void CreatePdfFromList(List<MaterialItem> Items, string filePath)
        {
            
            // Create a document object
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filePath, FileMode.Create));
           
            // Open the document for writing
            document.Open();

            // Add a table
            PdfPTable table = new PdfPTable(6); // 6 columns
            
            List<List<MaterialItem>> materialItems = new List<List<MaterialItem>>();
            List<MaterialItem> subMaterial  = new List<MaterialItem>();
            string hold = string.Empty;
            MaterialItem material1 = null;
            bool flag = false;
            double Total_length = 0;
            double total_area = 0;
            double total_weight = 0;
            Items.Add(new MaterialItem("","","","","",""));
            // Add headers
            table.AddCell("Profile");
            table.AddCell("Grade");
            table.AddCell("Qty");
            table.AddCell("Length");
            table.AddCell("Area");
            table.AddCell("Weight(lbs)");
            
            for (int i = 1; i < Items.Count; i++)
            {
                Total_length += ParseMeasurementToInches(Items[i - 1].Length);
                total_area += double.Parse(Items[i - 1].Area_in2);
                total_weight += double.Parse(Items[i - 1].Weight) * double.Parse(Items[i - 1].Qty);
                if (Items[i-1].Profile != Items[i].Profile)
                {
                    table.AddCell(Items[i-1].Profile);
                    table.AddCell(Items[i - 1].Grade);
                    table.AddCell(Items[i - 1].Qty);
                    table.AddCell(Items[i - 1].Length);
                    table.AddCell(Items[i - 1].Area_in2);
                    table.AddCell(Items[i - 1].Weight);
                    //Add headers
                    table.AddCell("total");
                    table.AddCell("");
                    table.AddCell("");
                    table.AddCell(ConvertInchesToFeetAndInches(Total_length));
                    table.AddCell(total_area.ToString());
                    table.AddCell(total_weight.ToString());
                    for (int j = 0; j < 6; j++)
                        table.AddCell("----------");
                    if((i+1) != Items.Count)
                    {
                        // Add headers
                        table.AddCell("Profile");
                        table.AddCell("Grade");
                        table.AddCell("Qty");
                        table.AddCell("Length");
                        table.AddCell("Area");
                        table.AddCell("Weight(lbs)");
                         Total_length = 0;
                         total_area = 0;
                         total_weight = 0;
                    }
                }
                else
                {
                    table.AddCell(Items[i - 1].Profile);
                    table.AddCell(Items[i - 1].Grade);
                    table.AddCell(Items[i - 1].Qty);
                    table.AddCell(Items[i - 1].Length);
                    table.AddCell(Items[i - 1].Area_in2);
                    table.AddCell(Items[i - 1].Weight);
                }
            }

           

           

            // Add the table to the document
            document.Add(table);

            // Close the document
            document.Close();
            writer.Close();

            Console.WriteLine($"PDF file '{filePath}' created successfully.");
        }

        

        private static double ParseMeasurementToInches(string measurement)
        {
            // Regular expressions to match the parts of the measurement
            string pattern = @"(?<feet>\d+)'\s*(?<inches>\d+)?\s*(?<fraction>\d+/\d+)?\""|(?<inchesOnly>\d+)?\s*(?<fractionOnly>\d+/\d+)?\""|(?<feetOnly>\d+)'";
            var match = Regex.Match(measurement, pattern);

            double inches = 0;
            double feet = 0;
            double fraction = 0;

            if (match.Success)
            {
                if (match.Groups["feet"].Success)
                {
                    feet = double.Parse(match.Groups["feet"].Value);
                }
                if (match.Groups["inches"].Success)
                {
                    inches = double.Parse(match.Groups["inches"].Value);
                }
                if (match.Groups["fraction"].Success)
                {
                    var fractionParts = match.Groups["fraction"].Value.Split('/');
                    fraction = double.Parse(fractionParts[0]) / double.Parse(fractionParts[1]);
                }
                if (match.Groups["inchesOnly"].Success)
                {
                    inches = double.Parse(match.Groups["inchesOnly"].Value);
                }
                if (match.Groups["fractionOnly"].Success)
                {
                    var fractionParts = match.Groups["fractionOnly"].Value.Split('/');
                    fraction = double.Parse(fractionParts[0]) / double.Parse(fractionParts[1]);
                }
                if (match.Groups["feetOnly"].Success)
                {
                    feet = double.Parse(match.Groups["feetOnly"].Value);
                }
            }

            return feet * 12 + inches + fraction;
        }

        private static string ConvertInchesToFeetAndInches(double totalInches)
        {
            int feet = (int)(totalInches / 12);
            double remainingInches = totalInches % 12;

            // Format fractional part if there is any
            string fractionalInches = FormatFractionalInches(remainingInches);

            if (feet > 0)
            {
                return $"{feet}'-{fractionalInches}\"";
            }
            else
            {
                return $"{fractionalInches}\"";
            }
        }

        private static string FormatFractionalInches(double inches)
        {
            // Determine the whole number part and the fractional part
            int wholeInches = (int)inches;
            double fractionPart = inches - wholeInches;

            string fractionalString = wholeInches.ToString();

            if (fractionPart > 0)
            {
                int denominator = 16; // Use 16 as the denominator for simplicity (can change to 32, 64, etc.)
                int numerator = (int)(fractionPart * denominator);

                while (numerator % 2 == 0 && numerator != 0)
                {
                    numerator /= 2;
                    denominator /= 2;
                }

                fractionalString += $" {numerator}/{denominator}";
            }

            return fractionalString;
        }
    }

}
