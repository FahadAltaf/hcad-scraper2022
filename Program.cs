

using System.Web;
using HtmlAgilityPack;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

public class DataModelAPN
{
    public string SearchedAPN { get; set; }
    public string OwnerName { get; set; }
    //public string FirstName { get; set; }
    //public string LastName { get; set; }
    public string MailingAddress { get; set; }
    public string MailingCity { get; set; }
    public string MailingState { get; set; }
    public string MailingZip { get; set; }
    public string PropertyAddress { get; set; }
    public string PropertyCity { get; set; }
    public string PropertyState { get; set; }
    public string PropertyZip { get; set; }
    public string LegalDescription { get; set; }
    public string Type { get; set; }
    public string Foundation { get; set; }
    public string HeatingOrAc { get; set; }
    public string ExteriorWall { get; set; }
    public string Bathroom { get; set; }
    public string Bedroom { get; set; }
    public string TotalRooms { get; set; }
    public string Style { get; set; }
    public string YearBuilt { get; set; }
    public string EffectiveDate { get; set; }
    public string TotalLivingArea { get; set; }
    public string ErrorMessage { get; set; }
}
public class Program
{
    public static void Main()
    {
        List<string> apns = new List<string> {
"0010010000013",
"0010020000001",
"0010020000003",
"0010020000004",
"0010020000013",
"0010020000015",
"0010020000016",
"0010020000023",
"0010020000024",
"0010030000001",
"0010030000008",
"0010030000015",
"0010030000016",
"0010030000017",
"0010040000001",
"0010040000004",
"0010050000004",
"0010050000020",
"0010050000020",
"0010060000010",
"0010060000010",
"0010070000013",
"0010070000014",
"0010070000017",
"0010070000017",
"0010080000001",
"0010080000002",
"0010080000004",
"0010080000006",
"0010080000014",
"0010080000014",
"0010080000015",
"0010080000015",
"0010080000016",
"0010080000016",
"0010090000001",
"0010100000001",
"0010100000004",
"0010100000008",
"0010100000011",
"0010100000012",
"0010110000010",
"0010120000010",
"0010130000012",
"0010130000012",
"0010140000001",
"0010150000001",
"0010150000002",
"0010150000003",
"0010150000016",
"0010160000001",
"0010160000006",
"0010160000007",
"0010160000008",
"0010160000009",
"0010160000011",
"0010160000012",
"0010160000013",
"0010160000014",
"0010160000017",
"0010160000020",
"0010160000021",
"0010170000001",
"0010170000013",
"0010170000014",
"0010170000015",
"0010180000001",
"0010180000001",
"0010180000004",
"0010180000005",
"0010190000001",
"0010190000002",
"0010190000003",
"0010190000004",
"0010190000005",
"0010190000006",
"0010190000007",
"0010190000008",
"0010190000011",
"0010190000013",
"0010190000020",
"0010190000021",
"0010190000022",
"0010190000026",
"0010190000027",
"0010190000027",
"0010190000028",
"0010200000001",
"0010200000002",
"0010200000003",
"0010200000004",
"0010210000001",
"0010220000002",
"0010220000033",
"0010240000001",
"0010240000001",
"0010250000009",
"0010250000009",
"0010250000009",
"0010250000013",
        };
        ScrapeByAPN(apns);

    }

    private static bool SplitName(string name, out string fn, out string ln, out string fn1, out string ln1)
    {
        string firstName = string.Empty, lastName = string.Empty, firstName1 = string.Empty, lastName1 = string.Empty;
        bool converted = false;
        if (!string.IsNullOrEmpty(name))
        {
            try
            {
                if (name.Contains("&"))
                {
                    var pieces = name.Split('&');
                    if (pieces.Length == 2)
                    {
                        var parts = pieces[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).Where(x => x.Length > 2).ToArray();
                        switch (parts.Length)
                        {
                            case 2:
                                firstName = parts[1];
                                lastName = parts[0];
                                converted = true;

                                break;
                            case 3:
                                firstName = parts[1];
                                lastName = parts[0];
                                converted = true;

                                break;
                            default:
                                break;
                        }

                        var parts1 = pieces[1].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).Where(x => x.Length > 2).ToArray();
                        switch (parts1.Length)
                        {
                            case 1:
                                firstName1 = parts1[0];
                                lastName1 = parts[0];
                                converted = true;
                                break;
                            case 2:
                                firstName1 = parts1[1];
                                lastName1 = parts1[0];
                                converted = true;
                                break;
                            case 3:
                                firstName = parts[1];
                                lastName = parts[0];
                                converted = true;

                                break;
                            default:
                                break;
                        }
                    }
                }
                else
                {
                    var parts = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries).Where(x => x.Length > 2).ToArray();
                    switch (parts.Length)
                    {
                        case 2:
                            firstName = parts[1];
                            lastName = parts[0];
                            converted = true;
                            break;
                        case 3:
                            firstName = parts[1];
                            lastName = parts[0];
                            converted = true;

                            break;
                        case 4:
                            firstName = parts[1];
                            lastName = parts[0];
                            firstName1 = parts[3];
                            lastName1 = parts[2];
                            converted = true;
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to split name \"{0}\". Reason: {1}", name, ex.Message);
            }
        }

        fn = firstName;
        ln = lastName;
        fn1 = firstName1;
        ln1 = lastName1;
        return converted;
    }

    private static void ScrapeByAPN(List<string> apns)
    {
        try
        {
            List<DataModelAPN> entries = new List<DataModelAPN>();
            ChromeOptions options = new ChromeOptions();
            options.AddArguments((IEnumerable<string>)new List<string>()
            {
                 "--silent-launch",
                 "--no-startup-window",
                 "no-sandbox",
                 "headless",
            });

            ChromeDriverService defaultService = ChromeDriverService.CreateDefaultService();
            defaultService.HideCommandPromptWindow = true;

            using (var driver = new ChromeDriver(defaultService, options))
            {
                foreach (var apn in apns)
                {
                    DataModelAPN model = new DataModelAPN() { SearchedAPN = apn };
                    try
                    {
                        var url = "https://public.hcad.org/records/Real.asp";
                        driver.Navigate().GoToUrl(url);

                        IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                        wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));

                        driver.FindElement(By.Id("acct")).SendKeys(apn);
                        driver.FindElement(By.XPath("//*[@id=\"Real_acct\"]/table/tbody/tr[3]/td[3]/nobr/input[1]")).Click();
                        int waitCount = 0;
                        do
                        {
                            try
                            {
                                driver.SwitchTo().Frame(driver.FindElement(By.Id("quickframe")));
                            }
                            catch { }
                            Thread.Sleep(500);
                            waitCount += 1;
                        } while (waitCount < 60 && !(driver.PageSource.Contains("Ownership History") || (driver.PageSource.Contains("tax year :") && driver.PageSource.Contains("record(s).")) || driver.PageSource.Contains("Currently, there are NO") || driver.PageSource.Contains("Please enter additional search criteria to reduce the number of records returned.")));

                        if (waitCount >= 60)
                        {
                            Console.WriteLine("Waited too long. but nothing found");
                        }

                        if (driver.PageSource.Contains("Tax Year: "))
                        {
                            HtmlDocument doc = new HtmlDocument();
                            doc.LoadHtml(driver.PageSource);
                            var ownerNameAddressNode = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[5]/tbody/tr[2]/td[1]/table/tbody/tr/th");
                            if (ownerNameAddressNode != null)
                            {
                                var pieces = ownerNameAddressNode.InnerHtml.Split(new string[] { "<br>" }, StringSplitOptions.RemoveEmptyEntries);
                                if (pieces.Length == 3)
                                {
                                    var sub = new HtmlDocument();
                                    sub.LoadHtml(pieces[0]);

                                    var name = HttpUtility.HtmlDecode(sub.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());
                                    Console.WriteLine("Owner Name: {0}", name);
                                    model.OwnerName = name;

                                    var sub1 = new HtmlDocument();
                                    sub1.LoadHtml(pieces[1]);

                                    var staddress = sub1.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    Console.WriteLine("Street Address: {0}", staddress);
                                    model.MailingAddress = staddress;
                                }
                                else if (pieces.Length == 4)
                                {
                                    var sub = new HtmlDocument();
                                    sub.LoadHtml(pieces[0]);

                                    var name = HttpUtility.HtmlDecode(sub.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());
                                    Console.WriteLine("Owner Name: {0}", name);
                                    model.OwnerName = name;

                                    var sub1 = new HtmlDocument();
                                    sub1.LoadHtml(pieces[1]);

                                    var staddress = sub1.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    Console.WriteLine("Street Address: {0}", staddress);
                                    model.MailingAddress = staddress;

                                    var sub2 = new HtmlDocument();
                                    sub2.LoadHtml(pieces[2]);

                                    var citystatezip = sub2.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    var cszParts = citystatezip.Split(new string[] { "&nbsp;" }, StringSplitOptions.RemoveEmptyEntries);
                                    if (cszParts.Length == 3)
                                    {
                                        Console.WriteLine("City: {0}", cszParts[0]);
                                        Console.WriteLine("State: {0}", cszParts[1]);
                                        Console.WriteLine("Zip: {0}", cszParts[2]);
                                        model.MailingCity = cszParts[0];
                                        model.MailingState = cszParts[1];
                                        model.MailingZip = cszParts[2];
                                    }
                                }
                                else if (pieces.Length == 5)
                                {
                                    var sub = new HtmlDocument();
                                    sub.LoadHtml(pieces[0]);

                                    var name = HttpUtility.HtmlDecode(sub.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());

                                    model.OwnerName = name;

                                    var sub4 = new HtmlDocument();
                                    sub4.LoadHtml(pieces[1]);

                                    var name4 = HttpUtility.HtmlDecode(sub4.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());
                                    model.OwnerName += " " + name4;
                                    Console.WriteLine("Owner Name: {0}", model.OwnerName);


                                    var sub1 = new HtmlDocument();
                                    sub1.LoadHtml(pieces[2]);

                                    var staddress = sub1.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    Console.WriteLine("Street Address: {0}", staddress);
                                    model.MailingAddress = staddress;

                                    var sub2 = new HtmlDocument();
                                    sub2.LoadHtml(pieces[3]);

                                    var citystatezip = sub2.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    var cszParts = citystatezip.Split(new string[] { "&nbsp;" }, StringSplitOptions.RemoveEmptyEntries);
                                    if (cszParts.Length == 3)
                                    {
                                        Console.WriteLine("City: {0}", cszParts[0]);
                                        Console.WriteLine("State: {0}", cszParts[1]);
                                        Console.WriteLine("Zip: {0}", cszParts[2]);
                                        model.MailingCity = cszParts[0];
                                        model.MailingState = cszParts[1];
                                        model.MailingZip = cszParts[2];
                                    }
                                }
                                else if (pieces.Length == 6)
                                {
                                    var sub = new HtmlDocument();
                                    sub.LoadHtml(pieces[0]);

                                    var name = HttpUtility.HtmlDecode(sub.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());

                                    model.OwnerName = name;

                                    var subx = new HtmlDocument();
                                    subx.LoadHtml(pieces[1]);

                                    var namex = HttpUtility.HtmlDecode(subx.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());

                                    model.OwnerName += " " + namex;

                                    var sub4 = new HtmlDocument();
                                    sub4.LoadHtml(pieces[2]);

                                    var name4 = HttpUtility.HtmlDecode(sub4.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim());
                                    model.OwnerName += " " + name4;
                                    Console.WriteLine("Owner Name: {0}", model.OwnerName);

                                    var sub1 = new HtmlDocument();
                                    sub1.LoadHtml(pieces[3]);

                                    var staddress = sub1.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    Console.WriteLine("Street Address: {0}", staddress);
                                    model.MailingAddress = staddress;

                                    var sub2 = new HtmlDocument();
                                    sub2.LoadHtml(pieces[4]);

                                    var citystatezip = sub2.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    var cszParts = citystatezip.Split(new string[] { "&nbsp;" }, StringSplitOptions.RemoveEmptyEntries);
                                    if (cszParts.Length == 3)
                                    {
                                        Console.WriteLine("City: {0}", cszParts[0]);
                                        Console.WriteLine("State: {0}", cszParts[1]);
                                        Console.WriteLine("Zip: {0}", cszParts[2]);
                                        model.MailingCity = cszParts[0];
                                        model.MailingState = cszParts[1];
                                        model.MailingZip = cszParts[2];
                                    }
                                }

                               /* string fn, ln, fn1, ln1;
                                fn = ln = fn1 = ln1 = string.Empty;
                                if (SplitName(model.FullName, out fn, out ln, out fn1, out ln1))
                                {
                                    Console.Write("{0}\n{1}\n{2}\n{3}", fn, ln, fn1, ln1);
                                }
                                model.FirstName = fn;
                                model.LastName = ln;*/
                               // model.FirstName2 = fn1;
                               // model.LastName2 = ln1;
                            }

                            var AddressNode = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[5]/tbody/tr[2]/td[2]/table/tbody/tr[2]/th");
                            if (AddressNode != null)
                            {
                                var pieces = AddressNode.InnerHtml.Split(new string[] { "<br>" }, StringSplitOptions.RemoveEmptyEntries);
                                if (pieces.Length > 0)
                                {
                                    var sub2 = new HtmlDocument();
                                    sub2.LoadHtml(pieces.LastOrDefault());
                                    model.PropertyAddress = pieces.FirstOrDefault();
                                    var citystatezip = sub2.DocumentNode.InnerText.Replace("\n", "").Replace("\r", "").Trim();
                                    var cszParts = citystatezip.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                                    if (cszParts.Length == 3)
                                    {
                                        Console.WriteLine("City: {0}", cszParts[0]);
                                        Console.WriteLine("State: {0}", cszParts[1]);
                                        Console.WriteLine("Zip: {0}", cszParts[2]);
                                        model.PropertyCity = cszParts[0];
                                        model.PropertyState = cszParts[1];
                                        model.PropertyZip = cszParts[2];
                                    }
                                    else if (cszParts.Length == 4)
                                    {
                                        Console.WriteLine("City: {0} {1}", cszParts[0], cszParts[1]);
                                        Console.WriteLine("State: {0}", cszParts[2]);
                                        Console.WriteLine("Zip: {0}", cszParts[3]);
                                        model.PropertyCity = cszParts[0] + " " + cszParts[1];
                                        model.PropertyState = cszParts[2];
                                        model.PropertyZip = cszParts[3];
                                    }
                                    else if (cszParts.Length == 5)
                                    {
                                        Console.WriteLine("City: {0} {1} {2}", cszParts[0], cszParts[1], cszParts[2]);
                                        Console.WriteLine("State: {0}", cszParts[3]);
                                        Console.WriteLine("Zip: {0}", cszParts[4]);
                                        model.PropertyCity = cszParts[0] + " " + cszParts[1] + " " + cszParts[2];
                                        model.PropertyState = cszParts[3];
                                        model.PropertyZip = cszParts[4];
                                    }
                                }
                            }

                            var classCode = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[5]/tbody/tr[2]/td[2]/table/tbody/tr[1]/th");
                            if (classCode != null)
                            {
                                var code = HttpUtility.HtmlDecode(classCode.InnerText.Trim());
                                Console.WriteLine("Class Code: {0}", code);
                                model.LegalDescription = code;
                            }

                            var typeNode = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[15]/tbody/tr[3]/td[3]");
                            if (typeNode != null)
                            {
                                var code = HttpUtility.HtmlDecode(typeNode.InnerText.Trim());
                                Console.WriteLine("Type: {0}", code);
                                model.Type = code;
                            }

                            var styleNode = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[15]/tbody/tr[3]/td[4]");
                            if (styleNode != null)
                            {
                                var code = HttpUtility.HtmlDecode(styleNode.InnerText.Trim());
                                Console.WriteLine("Type: {0}", code);
                                model.Style = code;
                            }

                            var totalLivingArea = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[6]/tbody/tr[4]/td[2]");
                            if (totalLivingArea != null)
                            {
                                var code = HttpUtility.HtmlDecode(totalLivingArea.InnerText.Trim());
                                Console.WriteLine("Total Living Area: {0}", code);
                                model.TotalLivingArea = code;
                            }

                            var builtYear = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[15]/tbody/tr[3]/td[2]");
                            if (builtYear != null)
                            {
                                var code = HttpUtility.HtmlDecode(builtYear.InnerText.Trim());
                                Console.WriteLine("Year Built: {0}", code);
                                model.YearBuilt = code;
                            }
                                                                            
                            var section = doc.DocumentNode.SelectNodes("//table").LastOrDefault(x=>x.InnerText.Contains("Building Data"));
                            if (section != null)
                            {
                                foreach (var row in section.ChildNodes[1].ChildNodes.Where(x=>x.Name=="tr"))
                                {
                                    try
                                    {
                                        if (row.ChildNodes.Where(x=>x.Name=="td").Count() == 2)
                                        {
                                            HtmlDocument sub = new HtmlDocument();
                                            sub.LoadHtml(row.InnerHtml);

                                            var header = sub.DocumentNode.SelectSingleNode("/td[1]").InnerText;
                                            var value = HttpUtility.HtmlDecode(sub.DocumentNode.SelectSingleNode("/td[2]").InnerText);
                                            switch (header)
                                            {
                                                case "Foundation Type":
                                                    model.Foundation = value;
                                                    break;
                                                case "Heating / AC":
                                                    model.HeatingOrAc = value;
                                                    break;
                                                case "Exterior Wall":
                                                    model.ExteriorWall = value;
                                                    break;
                                                case "Room:  Total":
                                                    model.TotalRooms = value;
                                                    break;
                                                case "Room:  Full Bath":
                                                    model.Bathroom = value;
                                                    break;
                                                case "Room:  Bedroom":
                                                    model.Bedroom = value;
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                            }

                            var link = doc.DocumentNode.SelectSingleNode("/html/body/table/tbody/tr/td/table[4]/tbody/tr/td/a");
                            if (link != null)
                            {
                                try
                                {
                                    var lk = "https://public.hcad.org" + link.Attributes.FirstOrDefault(x => x.Name == "href").Value;
                                    driver.Navigate().GoToUrl(lk);
                                    wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
                                    model.EffectiveDate = driver.FindElement(By.XPath("/html/body/table[2]/tbody/tr[4]/td[2]")).Text;
                                }
                                catch (Exception ex)
                                {

                                }
                                
                            }
                        }
                        entries.Add(model);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unable to process \"" + apn + "\". Reason: " + ex.Message);
                        // message = "Unable to process \"" + address + "\". Reason: " + ex.Message;
                    }
                }
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage())
            {
                var time = DateTime.Now.ToString("yyyyMMddHHmmss");
                var worksheet = excel.Workbook.Worksheets.Add($"Data");
                worksheet.Cells.LoadFromCollection(entries, true);
                excel.SaveAs($"result.xlsx");
            }

            Console.WriteLine("Done");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Unable to continue. Reason: " + ex.Message);
        }

    }
}

