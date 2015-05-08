using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DiversityInfo
{
    class Extractor
    {
        static Regex question = new Regex("([0-9]+)\\.");

        static List<string> headers = new List<string> { "Male", "Female", "PNTS", "Under 29", "30 to 39", "40 to 49", "50 to 59", "60 to 64", "65+", "PNTS", "Bangladeshi", "Chinese", "Indian", "Pakistani", "Any other Asian background", "African", "Caribbean", "Any other Black/African/Caribbean background", "White and Asian", "White and Black African", "White and Black Caribbean", "Any other mixed background", "Arab", "Any other ethnic group", "White", "PNTS", "Yes", "No", "PNTS", "Heterosexual/straight", "Gay /Lesbian", "Bisexual", "Other", "PNTS", "No religion", "Buddhist", "Christian", "Hindu", "Jewish", "Muslim", "Sikh", "Any other religion", "PNTS", "Full-time", "Part-time", "Job Share", "Other", "PNTS", "None", "Primary carer of a child/children (under 18)", "Primary carer of disabled child/children", "Primary carer of disabled adult (18 and over)", "Primary carer of older person (65 and over)", "Secondary carer", "PNTS", "Home Dept", "OGD", "Public service", "Voluntary", "Private Sector", "Other", "PNTS", "Yes", "No", "PNTS", "FLS", "HPDS", "SLS", "Other", "None", "PNTS", "CS Employee", "CS Jobs", "Guardian", "Exec/FT", "LinkedIn", "TimesOnline", "Twitter", "Word of Mouth", "Other", "PNTS" };

        internal static void Go(string folder, Func<string, int> stdout)
        {
            stdout("Starting");

            var csv = new StringBuilder();

            csv.AppendLine(string.Join(",", headers));

            foreach (var filename in Directory.GetFiles(folder, "*.docx"))
            {
                stdout("Processing: " + filename);
                var result = ExtractFile(filename);
                csv.AppendLine(string.Join(",", result));
            }

            string outfile = Path.Combine(folder, "results.csv");

            stdout("Writing to " + outfile);

            File.WriteAllText(outfile, csv.ToString());

            stdout("Finished");

            System.Console.Read();
        }

        private static List<string> ExtractFile(string filename)
        {
            List<Dictionary<string, string>> answers = new List<Dictionary<string, string>>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
            {
                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                System.Console.WriteLine(table);

                Dictionary<string, string> currentQuestion = null;

                foreach (TableRow row in table.Elements<TableRow>())
                {
                    TableCell firstCell = row.Elements<TableCell>().First();

                    Match m = question.Match(firstCell.InnerText);
                    if (m.Success)
                    {
                        currentQuestion = new Dictionary<string, string>();
                        answers.Add(currentQuestion);
                    }
                    else
                    {
                        string currentKey = null;
                        foreach (TableCell cell in row.Elements<TableCell>())
                        {
                            if (currentKey == null)
                            {
                                currentKey = cell.InnerText.Trim();
                            }
                            else if (currentKey == "")
                            {
                                currentKey = null;
                            }
                            else
                            {
                                CheckBox cb = cell.Descendants<CheckBox>().First();
                                Checked state = cb.GetFirstChild<Checked>();
                                if (state != null && state.Val == null)
                                {
                                    currentQuestion.Add(currentKey, "1");
                                }
                                else
                                {
                                    currentQuestion.Add(currentKey, "0");
                                }

                                currentKey = null;
                            }
                        }
                    }
                }

                List<string> result = new List<string>();
                result.Add(answers[0]["Male"]);
                result.Add(answers[0]["Female"]);
                result.Add(answers[0]["Prefer not to say"]);

                result.Add(answers[1]["Under 29"]);
                result.Add(answers[1]["30 to 39"]);
                result.Add(answers[1]["40 to 49"]);
                result.Add(answers[1]["50 to 59"]);
                result.Add(answers[1]["60 to 64"]);
                result.Add(answers[1]["65 and over"]);
                result.Add(answers[1]["Prefer not to say"]);

                result.Add(answers[2]["Bangladeshi"]);
                result.Add(answers[2]["Chinese"]);
                result.Add(answers[2]["Indian"]);
                result.Add(answers[2]["Pakistani"]);
                result.Add(answers[2]["Any other Asian background"]);
                result.Add(answers[2]["African"]);
                result.Add(answers[2]["Caribbean"]);
                result.Add(answers[2]["Any other Black/African/Caribbean background"]);
                result.Add(answers[2]["White and Asian"]);
                result.Add(answers[2]["White and Black African"]);
                result.Add(answers[2]["White and Black Caribbean"]);
                result.Add(answers[2]["Any other mixed  / multiple ethnic background"]);
                result.Add(answers[2]["Arab"]);
                result.Add(answers[2]["Any other ethnic group"]);
                result.Add(answers[2]["White"]);
                result.Add(answers[2]["Prefer not to say"]);

                result.Add(answers[3]["Yes"]);
                result.Add(answers[3]["No"]);
                result.Add(answers[3]["Prefer not to say"]);

                result.Add(answers[4]["Heterosexual / Straight"]);
                result.Add(answers[4]["Gay / Lesbian"]);
                result.Add(answers[4]["Bisexual"]);
                result.Add(answers[4]["Other"]);
                result.Add(answers[4]["Prefer not say"]);

                result.Add(answers[5]["No religion"]);
                result.Add(answers[5]["Buddhist"]);
                result.Add(answers[5]["Christian"]);
                result.Add(answers[5]["Hindu"]);
                result.Add(answers[5]["Jewish"]);
                result.Add(answers[5]["Muslim"]);
                result.Add(answers[5]["Sikh"]);
                result.Add(answers[5]["Any other religion"]);
                result.Add(answers[5]["Prefer not to say"]);

                result.Add(answers[6]["Full-time"]);
                result.Add(answers[6]["Part-time"]);
                result.Add(answers[6]["Job Share"]);
                result.Add(answers[6]["Other"]);
                result.Add(answers[6]["Prefer not to say"]);

                result.Add(answers[7]["None"]);
                result.Add(answers[7]["Primary carer of a child/children (under 18)"]);
                result.Add(answers[7]["Primary carer of disabled child/children"]);
                result.Add(answers[7]["Primary carer of disabled adult (18 and over)"]);
                result.Add(answers[7]["Primary carer of older person (65 and over)"]);
                result.Add(answers[7]["Secondary carer"]);
                result.Add(answers[7]["Prefer not to say"]);

                result.Add(answers[8]["Home Department of vacancy"]);
                result.Add(answers[8]["Other Government Dept."]);
                result.Add(answers[8]["Wider Public Service"]);
                result.Add(answers[8]["Voluntary Sector"]);
                result.Add(answers[8]["Private Sector"]);
                result.Add(answers[8]["Other"]);
                result.Add(answers[8]["Prefer not to say"]);

                result.Add(answers[9]["Yes"]);
                result.Add(answers[9]["No"]);
                result.Add(answers[9]["Prefer not to say"]);

                result.Add(answers[10]["Future Leaders Scheme"]);
                result.Add(answers[10]["High Potential  Development Scheme"]);
                result.Add(answers[10]["Senior Leaders Scheme"]);
                result.Add(answers[10]["Other"]);
                result.Add(answers[10]["None"]);
                result.Add(answers[10]["Prefer not to say"]);

                result.Add(answers[11]["From a Civil Service employee"]);
                result.Add(answers[11]["From the Civil Service Jobs website"]);
                result.Add(answers[11]["Guardian Jobs"]);
                result.Add(answers[11]["Executive Appointments / Financial Times"]);
                result.Add(answers[11]["LinkedIn"]);
                result.Add(answers[11]["TimesOnline"]);
                result.Add(answers[11]["Twitter"]);
                result.Add(answers[11]["Word of Mouth"]);
                result.Add(answers[11]["Other"]);
                result.Add(answers[11]["Prefer not to say"]);

                return result;

            }
        }
    }

}
