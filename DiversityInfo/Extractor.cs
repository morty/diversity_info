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
    class MyKeyNotFoundException: KeyNotFoundException
    {
        private string key;
        public MyKeyNotFoundException(string key)
            : base("Key not found: '" + key + "'")
        {
            this.key = key;
        }
    }

    class Extractor
    {
        static Regex question = new Regex("([0-9]+)\\.");

        static string dept_code_validate = ".*";
        static string org_code_validate = "AGO|CPS|CPSI|SFO|TSOL|CO|CCS|CC|CMA|BIS|ACAS|CH|INS|LR|MO|NMO|OS|SFA|UKIPO|SA|DCLG|PI|QEII|DCMS|RP|DFE|EFA|NCTL|STA|DEFRA|APHA|CEFAS|FERA|RPA|VMD|OFWAT|DFID|DFT|DVLA|DSA|HA|MCA|ORR|VOSA|VCA|DWP|HSE|DECC|DH|FSA|MHPRA|PHE|ESTYN|FCO|FCOS|WP|HMRC|VO|HMT|DMO|GAD|NSI|HO|NCA|MOD|DSTL|DES|DSG|UKHO|MOJ|HMCTS|LAA|NA|NOMS|OPG|CICA|NIO|OFSTED|OFGEM|OFQAL|SO|SG|DS|SAA|SPPA|OAB|COPFS|SPS|OSCR|SCS|HS|NRS|RS|SHR|TS|ES|SIS|UKSA|UKEF";
        static string pay_band_validate = ".*";
        static string status_code_validate = "CW|US|UI|S";
        static string prof_code_validate = "CM|EC|ENG|FIN|HR|IT|IA|LAW|KIM|MED|OPDEL|OR|PLA|POL|PCM|PPM|PSY|INS|SCI|SMR|STA|TAX|VET|PAM|OTH|NK";
        static string capacity_code_validate = "C|D|P|N";
        static string date_of_completion_validate = "[0-9]{6}";
        static string exception_validate = ".*";

        static List<string> headers = new List<string> { "Male", "Female", "PNTS", "Under 29", "30 to 39", "40 to 49", "50 to 59", "60 to 64", "65+", "PNTS", "Bangladeshi", "Chinese", "Indian", "Pakistani", "Any other Asian background", "African", "Caribbean", "Any other Black/African/Caribbean background", "White and Asian", "White and Black African", "White and Black Caribbean", "Any other mixed background", "Arab", "Any other ethnic group", "White", "PNTS", "Yes", "No", "PNTS", "Heterosexual/straight", "Gay /Lesbian", "Bisexual", "Other", "PNTS", "No religion", "Buddhist", "Christian", "Hindu", "Jewish", "Muslim", "Sikh", "Any other religion", "PNTS", "Full-time", "Part-time", "Job Share", "Other", "PNTS", "None", "Primary carer of a child/children (under 18)", "Primary carer of disabled child/children", "Primary carer of disabled adult (18 and over)", "Primary carer of older person (65 and over)", "Secondary carer", "PNTS", "Home Dept", "OGD", "Public service", "Voluntary", "Private Sector", "Other", "PNTS", "Yes", "No", "PNTS", "FLS", "HPDS", "SLS", "Other", "None", "PNTS", "CS Employee", "CS Jobs", "Guardian", "Exec/FT", "LinkedIn", "TimesOnline", "Twitter", "Word of Mouth", "Other", "PNTS", "Dept Ref Code", "Org Ref Code", "Vacancy Ref Number", "Pay Band", "Status Ref Code", "Professional Ref Code", "Key Capacity Ref Code", "Date Of Campaign Completion", "Exception 1", "Exception 2", "Exception 3"};

        internal static void Go(string folder, Func<string, int> stdout)
        {
            stdout("Starting");

            var csv = new StringBuilder();

            csv.AppendLine(string.Join(",", headers.ToArray()));

            HashSet<string> failed = new HashSet<string>();

            foreach (var filename in Directory.GetFiles(folder, "*.docx"))
            {
                stdout("\nProcessing: " + filename);
                List<string> result;
                try
                {
                     result = ExtractFile(filename, stdout, failed);
                }
                catch (Exception ex)
                {
                    stdout("------------------------------------------");
                    stdout("Error processing file: " + filename);
                    stdout(ex.Message);
                    stdout(ex.StackTrace);
                    stdout("------------------------------------------");
                    failed.Add(filename);
                    continue;
                }
                csv.AppendLine(string.Join(",", result.ToArray()));
            }

            string outfile = Path.Combine(folder, "results.csv");


            try
            {
                File.WriteAllText(outfile, csv.ToString());
            }
            catch (IOException)
            {
                stdout("Error writing to output file. Please check that this is not already opened by another program (e.g. Excel).");
                return;
            }
            

            stdout("Finished\n");
            if (failed.Count() > 0)
            {
                stdout("Number of errors: " + failed.Count() + "\n");
                stdout("Failed files:");
                foreach (var filename in failed)
                {
                    stdout(filename);
                }
            }
        }

        private static List<string> ExtractFile(string filename, Func<string, int> stdout, HashSet<string> failed)
        {
            List<Dictionary<string, string>> answers = new List<Dictionary<string, string>>();

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filename, true))
            {
                Table table1 = doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                Dictionary<string, string> currentQuestion = null;

                foreach (TableRow row in table1.Elements<TableRow>())
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
                result.Add(getValue(answers[0], "Male"));
                result.Add(getValue(answers[0], "Female"));
                result.Add(getValue(answers[0], "Prefer not to say"));

                result.Add(getValue(answers[1], "29 or under"));
                result.Add(getValue(answers[1], "30 to 39"));
                result.Add(getValue(answers[1], "40 to 49"));
                result.Add(getValue(answers[1], "50 to 59"));
                result.Add(getValue(answers[1], "60 to 64"));
                result.Add(getValue(answers[1], "65 and over"));
                result.Add(getValue(answers[1], "Prefer not to say"));

                result.Add(getValue(answers[2], "Bangladeshi"));
                result.Add(getValue(answers[2], "Chinese"));
                result.Add(getValue(answers[2], "Indian"));
                result.Add(getValue(answers[2], "Pakistani"));
                result.Add(getValue(answers[2], "Any other Asian background"));
                result.Add(getValue(answers[2], "African"));
                result.Add(getValue(answers[2], "Caribbean"));
                result.Add(getValue(answers[2], "Any other Black/African/Caribbean background"));
                result.Add(getValue(answers[2], "White and Asian"));
                result.Add(getValue(answers[2], "White and Black African"));
                result.Add(getValue(answers[2], "White and Black Caribbean"));
                result.Add(getValue(answers[2], "Any other mixed  / multiple ethnic background"));
                result.Add(getValue(answers[2], "Arab"));
                result.Add(getValue(answers[2], "Any other ethnic group"));
                result.Add(getValue(answers[2], "White"));
                result.Add(getValue(answers[2], "Prefer not to say"));

                result.Add(getValue(answers[3], "Yes"));
                result.Add(getValue(answers[3], "No"));
                result.Add(getValue(answers[3], "Prefer not to say"));

                result.Add(getValue(answers[4], "Heterosexual / Straight"));
                result.Add(getValue(answers[4], "Gay / Lesbian"));
                result.Add(getValue(answers[4], "Bisexual"));
                result.Add(getValue(answers[4], "Other"));
                result.Add(getValue(answers[4], "Prefer not say"));

                result.Add(getValue(answers[5], "No religion"));
                result.Add(getValue(answers[5], "Buddhist"));
                result.Add(getValue(answers[5], "Christian"));
                result.Add(getValue(answers[5], "Hindu"));
                result.Add(getValue(answers[5], "Jewish"));
                result.Add(getValue(answers[5], "Muslim"));
                result.Add(getValue(answers[5], "Sikh"));
                result.Add(getValue(answers[5], "Any other religion"));
                result.Add(getValue(answers[5], "Prefer not to say"));

                result.Add(getValue(answers[6], "Full-time"));
                result.Add(getValue(answers[6], "Part-time"));
                result.Add(getValue(answers[6], "Job Share"));
                result.Add(getValue(answers[6], "Other"));
                result.Add(getValue(answers[6], "Prefer not to say"));

                result.Add(getValue(answers[7], "None"));
                result.Add(getValue(answers[7], "Primary carer of a child/children (under 18)"));
                result.Add(getValue(answers[7], "Primary carer of disabled child/children"));
                result.Add(getValue(answers[7], "Primary carer of disabled adult (18 and over)"));
                result.Add(getValue(answers[7], "Primary carer of older person (65 and over)"));
                result.Add(getValue(answers[7], "Secondary carer"));
                result.Add(getValue(answers[7], "Prefer not to say"));

                result.Add(getValue(answers[8], "Home department of vacancy"));
                result.Add(getValue(answers[8], "Other government dept."));
                result.Add(getValue(answers[8], "Wider Public Service"));
                result.Add(getValue(answers[8], "Voluntary Sector"));
                result.Add(getValue(answers[8], "Private Sector"));
                result.Add(getValue(answers[8], "Other"));
                result.Add(getValue(answers[8], "Prefer not to say"));

                result.Add(getValue(answers[9], "Yes"));
                result.Add(getValue(answers[9], "No"));
                result.Add(getValue(answers[9], "Prefer not to say"));

                result.Add(getValue(answers[10], "Future Leaders Scheme"));
                result.Add(getValue(answers[10], "High Potential  Development Scheme"));
                result.Add(getValue(answers[10], "Senior Leaders Scheme"));
                result.Add(getValue(answers[10], "Other"));
                result.Add(getValue(answers[10], "None"));
                result.Add(getValue(answers[10], "Prefer not to say"));

                result.Add(getValue(answers[11], "From a Civil Service employee"));
                result.Add(getValue(answers[11], "From the Civil Service Jobs website"));
                result.Add(getValue(answers[11], "Guardian Jobs"));
                result.Add(getValue(answers[11], "Executive Appointments / Financial Times"));
                result.Add(getValue(answers[11], "LinkedIn"));
                result.Add(getValue(answers[11], "TimesOnline"));
                result.Add(getValue(answers[11], "Twitter"));
                result.Add(getValue(answers[11], "Word of Mouth"));
                result.Add(getValue(answers[11], "Other"));
                result.Add(getValue(answers[11], "Prefer not to say"));

                Table table2 = doc.MainDocumentPart.Document.Body.Elements<Table>().ElementAt(1);
                var rows = table2.Elements<TableRow>();

                var dept_code = getTableValue(rows, 1, 1).Trim().ToUpper();
                if (!Regex.IsMatch(dept_code, dept_code_validate))
                {
                    //throw new Exception("Dept Code is not valid: " + dept_code);
                    stdout("Dept Code is not valid: " + dept_code);
                    failed.Add(filename);
                }
                result.Add(dept_code);

                var org_code = getTableValue(rows, 2, 1).Trim().ToUpper();
                if (!Regex.IsMatch(org_code, org_code_validate))
                {
                    //throw new Exception("Org Code is not valid: " + org_code);
                    stdout("Org Code is not valid: " + org_code);
                    failed.Add(filename);
                }
                result.Add(org_code);

                // Vacancy Reference Number not validated
                result.Add(getTableValue(rows, 3, 1));

                var pay_band = getTableValue(rows, 4, 1).Trim().ToUpper();
                if (!Regex.IsMatch(pay_band, pay_band_validate))
                {
                    //throw new Exception("Pay Band is not valid: " + pay_band);
                    stdout("Pay Band is not valid: " + pay_band);
                    failed.Add(filename);
                }
                result.Add(pay_band);

                var status_code = getTableValue(rows, 5, 1).Trim().ToUpper();
                if (!Regex.IsMatch(status_code, status_code_validate))
                {
                    //throw new Exception("Status Code is not valid: " + status_code);
                    stdout("Status Code is not valid: " + status_code);
                    failed.Add(filename);
                }
                result.Add(status_code);

                var prof_code = getTableValue(rows, 6, 1).Trim().ToUpper();
                if (!Regex.IsMatch(prof_code, prof_code_validate))
                {
                    //throw new Exception("Profession Code is not valid: " + prof_code);
                    stdout("Profession Code is not valid: " + prof_code);
                    failed.Add(filename);
                }
                result.Add(prof_code);

                var capacity_code = getTableValue(rows, 7, 1).Trim().ToUpper();
                if (!Regex.IsMatch(capacity_code, capacity_code_validate))
                {
                    //throw new Exception("Key Capacity Code is not valid: " + capacity_code);
                    stdout("Key Capacity Code is not valid: " + capacity_code);
                    failed.Add(filename);
                }
                result.Add(capacity_code);

                /*
                var date_of_completion = getTableValue(rows, 8, 1);
                if (!Regex.IsMatch(date_of_completion, date_of_completion_validate))
                {
                    //throw new Exception("Date Of Completion is not valid: " + date_of_completion);
                    stdout("Date Of Completion is not valid: " + date_of_completion);
                    failed.Add(filename);
                }
                result.Add(date_of_completion);
                */

                var exception1 = getTableValue(rows, 8, 1);
                if (!Regex.IsMatch(exception1, exception_validate))
                {
                    //throw new Exception("E1 is not valid: " + exception1);
                    stdout("E1 is not valid: " + exception1);
                    failed.Add(filename);
                }
                result.Add(exception1);

                var exception2 = getTableValue(rows, 9, 1);
                if (!Regex.IsMatch(exception2, exception_validate))
                {
                    //throw new Exception("E2 is not valid: " + exception2);
                    stdout("E2 is not valid: " + exception1);
                    failed.Add(filename);
                }
                result.Add(exception2);

                var exception3 = getTableValue(rows, 10, 1);
                if (!Regex.IsMatch(exception3, exception_validate))
                {
                    //throw new Exception("E3 is not valid: " + exception3);
                    stdout("E3 is not valid: " + exception1);
                    failed.Add(filename);
                }
                result.Add(exception3);

                return result;
            }
        }

        private static string getTableValue(IEnumerable<TableRow> rows, int i, int j)
        {
            return rows.ElementAt(i).Elements<TableCell>().ElementAt(j).InnerText.Trim();
        }

        private static string getValue(Dictionary<string, string> dictionary, string key)
        {
            try
            {
                return dictionary[key];
            }
            catch (KeyNotFoundException)
            {
                
                throw new MyKeyNotFoundException(key);
            }
        }
    }

}
