using OpenQA.Selenium;

using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClosedXML.Excel;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Chrome;

using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;

namespace EudamedAutomation
{
    class Program
    {
        static void Main(string[] args)
        {

            bool retryIteration = false;
            // Initialize WebDriver
            var options = new ChromeOptions();
            options.AddArgument("start-maximized"); // Maximizes the browser window
            options.AddArguments("--no-sandbox");
            options.AddArguments("--disable-dev-shm-usage");
            options.AddArguments("--remote-debugging-port=9222");
            options.AddArguments("--disable-gpu");
            options.AddArguments("--window-size=1920,1080");

            IWebDriver driver = new ChromeDriver(options);

            Console.WriteLine("Initializing Chrome WebDriver and maximizing the browser window...");
            int totalPages = 22222; // Total number of pages

            try
            {
                // Open the webpage
                Console.WriteLine("Navigating to the Eudamed website...");
                driver.Navigate().GoToUrl("https://ec.europa.eu/tools/eudamed/#/screen/search-device?submitted=true");

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(40));

                // Wait for the dropdown trigger element to be visible and click it
                Console.WriteLine("Waiting for the dropdown trigger to be visible...");

                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                // Wait until the button for the next page is present
                //var nextPageButton = wait.Until(d => d.FindElement(By.XPath(nextPageButtonXPath)));
                IWebElement dropdownTrigger = wait.Until(ExpectedConditions.ElementToBeClickable(By.ClassName("p-dropdown")));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", dropdownTrigger);
                Console.WriteLine("Clicking the dropdown to select '50 items per page'...");
                dropdownTrigger.Click();

                // Wait for the option with aria-label='50' to become visible
                Console.WriteLine("Waiting for the '50 items per page' option...");
                IWebElement dropdownOption = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("[aria-label='50']")));
                js.ExecuteScript("arguments[0].scrollIntoView(true);", dropdownOption);
                dropdownOption.Click();

                // Wait for the page to load with 50 items

                Console.WriteLine("Waiting for the page to load with 50 items per page...");
                wait.Until(d =>
                {
                    try
                    {
                        // Find the table or the specific section where the items are located
                        var table = d.FindElement(By.TagName("p-table")); // Update this selector to target the correct table element
                        var rows = table.FindElements(By.CssSelector("tbody > tr")); // Adjust to match the row selector

                        // Ensure the page is showing 50 items (rows) per page
                        return rows.Count == 50;
                    }
                    catch (NoSuchElementException)
                    {
                        return false; // Continue waiting if the table is not found
                    }
                });

                Console.WriteLine("The page with 50 items per page has loaded successfully.");

                // Wait for the table to stabilize
                Console.WriteLine("Waiting for the table to stabilize...");
                Thread.Sleep(5000); // Adjust the sleep time as needed based on the page load time

                // Define the sequence of pages to click
                int[] pagesToVisit = {5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 41, 43,
                45, 47, 49, 51, 53, 55, 57, 59, 61, 63, 65, 67, 69, 71, 73, 75, 77, 79, 81, 83,
                85, 87, 89, 91, 93, 95, 97, 99, 101, 103, 105, 107, 109, 111, 113, 115, 117, 119,
                121, 123, 125, 127, 129, 131, 133, 135, 137, 139, 141, 143, 145, 147, 149, 151,
                153, 155, 157, 159, 161, 163, 165, 167, 169, 171, 173, 175, 177, 179, 181, 183,
                185, 187, 189, 191, 193, 195, 197, 199, 201, 203, 205, 207, 209, 211, 213, 215,
                217, 219, 221, 223, 225, 227, 229, 231, 233, 235, 237, 239, 241, 243, 245, 247,
                249, 251, 253, 255, 257, 259, 261, 263, 265, 267, 269, 271, 273, 275, 277, 279,
                281, 283, 285, 287, 289, 291, 293, 295, 297, 299, 301, 303, 305, 307, 309, 311,
                313, 315, 317, 319, 321, 323, 325, 327, 329, 331, 333, 335, 337, 339, 341, 343,
                345, 347, 349, 351, 353, 355, 357, 359, 361, 363, 365, 367, 369, 371, 373, 375,
                377, 379, 381, 383, 385, 387, 389, 391, 393, 395, 397, 399, 401, 403, 405, 407,
                409, 411, 413, 415, 417, 419, 421, 423, 425, 427, 429, 431, 433, 435, 437, 439,
                441, 443, 445, 447, 449, 451, 453, 455, 457, 459, 461, 463, 465, 467, 469, 471,
                473, 475, 477, 479, 481, 483, 485, 487, 489, 491, 493, 495, 497, 499, 501, 503,
                505, 507, 509, 511, 513, 515, 517, 519, 521, 523, 525, 527, 529, 531, 533, 535,
                537, 539, 541, 543, 545, 547, 549, 551, 553, 555, 557, 559, 561, 563, 565, 567,
                569, 571, 573, 575, 577, 579, 581, 583, 585, 587, 589, 591, 593, 595, 597, 599,
                601, 603, 605, 607, 609, 611, 613, 615, 617, 619, 621, 623, 625, 627, 629, 631,
                633, 635, 637, 639, 641, 643, 645, 647, 649, 651, 653, 655, 657, 659, 661, 663,
                665, 667, 669, 671, 673, 675, 677, 679, 681, 683, 685, 687, 689, 691, 693, 695,
                697, 699, 701, 703, 705, 707, 709, 711, 713, 715, 717, 719, 721, 723, 725, 727,
                729, 731, 733, 735, 737, 739, 741, 743, 745, 747, 749, 751, 753, 755, 757, 759,
                761, 763, 765, 767, 769, 771, 773, 775, 777, 779, 781, 783, 785, 787, 789, 791,
                793, 795, 797, 799, 801, 803, 805, 807, 809, 811, 813, 815, 817, 819, 821, 823,
                825, 827, 829, 831, 833, 835, 837, 839, 841, 843, 845, 847, 849, 851, 853, 855,
                857, 859, 861, 863, 865, 867, 869, 871, 873, 875, 877, 879, 881, 883, 885, 887,
                889, 891, 893, 895, 897, 899, 901, 903, 905, 907, 909, 911, 913, 915, 917, 919,
                921, 923, 925, 927, 929, 931, 933, 935, 937, 939, 941, 943, 945, 947, 949, 951,
                953, 955, 957, 959, 961, 963, 965, 967, 969, 971, 973, 975, 977, 979, 981, 983,
                985, 987, 989, 991, 993, 995, 997, 999, 1001, 1003, 1005, 1007, 1009, 1011, 1013,
                1015, 1017, 1019, 1021, 1023, 1025, 1027, 1029, 1031, 1033, 1035, 1037, 1039, 1041,
                1043, 1045, 1047, 1049, 1051, 1053, 1055, 1057, 1059, 1061, 1063, 1065, 1067, 1069,
                1071, 1073, 1075, 1077, 1079, 1081, 1083, 1085, 1087, 1089, 1091, 1093, 1095, 1097,
                1099, 1101, 1103, 1105, 1107, 1109, 1111, 1113, 1115, 1117, 1119, 1121, 1123, 1125,
                1127, 1129, 1131, 1133, 1135, 1137, 1139, 1141, 1143, 1145, 1147, 1149, 1151, 1153,
                1155, 1157, 1159, 1161, 1163, 1165, 1167, 1169, 1171, 1173, 1175, 1177, 1179, 1181,
                1183, 1185, 1187, 1189, 1191, 1193, 1195, 1197, 1199, 1201, 1203, 1205, 1207, 1209,
                1211, 1213, 1215, 1217, 1219, 1221, 1223, 1225, 1227, 1229, 1231, 1233, 1235, 1237,
                1239, 1241, 1243, 1245, 1247, 1249, 1251, 1253, 1255, 1257, 1259, 1261, 1263, 1265,
                1267, 1269, 1271, 1273, 1275, 1277, 1279, 1281, 1283, 1285, 1287, 1289, 1291, 1293,
                1295, 1297, 1299, 1301, 1303, 1305, 1307, 1309, 1311, 1313, 1315, 1317, 1319, 1321,
                1323, 1325, 1327, 1329, 1331, 1333, 1335, 1337, 1339, 1341, 1343, 1345, 1347, 1349,
                1351, 1353, 1355, 1357, 1359, 1361, 1363, 1365, 1367, 1369, 1371, 1373, 1375, 1377,
                1379, 1381, 1383, 1385, 1387, 1389, 1391, 1393, 1395, 1397, 1399, 1401, 1403, 1405,
                1407, 1409, 1411, 1413, 1415, 1417, 1419, 1421, 1423, 1425, 1427, 1429, 1431, 1433,
                1435, 1437, 1439, 1441, 1443, 1445, 1447, 1449, 1451, 1453, 1455, 1457, 1459, 1461,
                1463, 1465, 1467, 1469, 1471, 1473, 1475, 1477, 1479, 1481, 1483, 1485, 1487, 1489,
                1491, 1493, 1495, 1497, 1499, 1501, 1503, 1505, 1507, 1509, 1511, 1513, 1515, 1517,
                1519, 1521, 1523, 1525, 1527, 1529, 1531, 1533, 1535, 1537, 1539, 1541, 1543, 1545, 
                1547, 1549, 1551, 1553, 1555, 1557, 1559, 1561, 1563, 1565, 1567, 1569, 1571, 1573,
                1575, 1577, 1579, 1581, 1583, 1585, 1587, 1589, 1591, 1593, 1595, 1597, 1599, 1601, 
                1603, 1605, 1607, 1609, 1611, 1613, 1615, 1617, 1619, 1621, 1623, 1625, 1627, 1629, 
                1631, 1633, 1635, 1637, 1639, 1641, 1643, 1645, 1647, 1649, 1651, 1653, 1655, 1657,
                1659, 1661, 1663, 1665, 1667, 1669, 1671, 1673, 1675, 1677, 1679, 1681, 1683, 1685, 
                1687, 1689, 1691, 1693, 1695, 1697, 1699, 1701, 1703, 1705, 1707, 1709, 1711, 1713, 
                1715, 1717, 1719, 1721, 1723, 1725, 1727, 1729, 1731, 1733, 1735, 1737, 1739, 1741, 
                1743, 1745, 1747, 1749, 1751, 1753, 1755, 1757, 1759, 1761, 1763, 1765, 1767, 1769, 
                1771, 1773, 1775, 1777, 1779, 1781, 1783, 1785, 1787, 1789, 1791, 1793, 1795, 1797, 
                1799, 1801, 1803, 1805, 1807, 1809, 1811, 1813, 1815, 1817, 1819, 1821, 1823, 1825, 
                1827, 1829, 1831, 1833, 1835, 1837, 1839, 1841, 1843, 1845, 1847, 1849, 1851, 1853, 
                1855, 1857, 1859, 1861, 1863, 1865, 1867, 1869, 1871, 1873, 1875, 1877, 1879, 1881, 
                1883, 1885, 1887, 1889, 1891, 1893, 1895, 1897, 1899, 1901, 1903, 1905, 1907, 1909, 
                1911, 1913, 1915, 1917, 1919, 1921, 1923, 1925, 1927, 1929, 1931, 1933, 1935, 1937, 
                1939, 1941, 1943, 1945, 1947, 1949, 1951, 1953, 1955, 1957, 1959, 1961, 1963, 1965, 
                1967, 1969, 1971, 1973, 1975, 1977, 1979, 1981, 1983, 1985, 1987, 1989, 1991, 1993, 
                1995, 1997, 1999, 2001, 2003, 2005, 2007, 2009, 2011, 2013, 2015, 2017, 2019, 2021, 
                2023, 2025, 2027, 2029, 2031, 2033, 2035, 2037, 2039, 2041, 2043, 2045, 2047, 2049, 
                2051, 2053, 2055, 2057, 2059, 2061, 2063, 2065, 2067, 2069, 2071, 2073, 2075, 2077, 
                2079, 2081, 2083, 2085, 2087, 2089, 2091, 2093, 2095, 2097, 2099, 2101, 2103, 2105, 
                2107, 2109, 2111, 2113, 2115, 2117, 2119, 2121, 2123, 2125, 2127, 2129, 2131, 2133, 
                2135, 2137, 2139, 2141, 2143, 2145, 2147, 2149, 2151, 2153, 2155, 2157, 2159, 2161,
                2163, 2165, 2167, 2169, 2171, 2173, 2175, 2177, 2179, 2181, 2183, 2185, 2187, 2189, 
                2191, 2193, 2195, 2197, 2199, 2201, 2203, 2205, 2207, 2208}; // Last page can be changed


                    foreach (int page in pagesToVisit)
                    {
                        

                        do
                        {
                            try
                            {
                                string pageXPath = $"//button[contains(@aria-label, 'Page number {page} ')]";
                                Console.WriteLine($"\nNavigating to Page {page}...");

                                IWebElement pageButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(pageXPath)));
                                js.ExecuteScript("arguments[0].scrollIntoView(true);", pageButton);
                                pageButton.Click();

                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"✅ Successfully navigated to Page {page}");
                                Console.ResetColor();

                                Thread.Sleep(3000);
                            }
                            catch (NoSuchElementException ex)
                            {
                                Console.WriteLine($"An error occurred: {ex.Message}");

                                // Prompt user for action
                                Console.WriteLine("Choose an option:");
                                Console.WriteLine("[R] Retry this iteration");
                                Console.WriteLine("[S] Skip to next page");
                                Console.WriteLine("[E] Exit program");

                                string userChoice = Console.ReadLine()?.ToUpper();
                                switch (userChoice)
                                {
                                    case "R":
                                        retryIteration = true; // Repeat this iteration
                                        Console.WriteLine("Retrying...");
                                        break;
                                    case "S":
                                        Console.WriteLine("Skipping to the next page...");
                                        retryIteration = false; // Moves to next iteration
                                        break;
                                    case "E":
                                        Console.WriteLine("Exiting program...");
                                        driver.Quit();
                                        return; // Exit function
                                    default:
                                        Console.WriteLine("Invalid choice. Skipping iteration.");
                                        retryIteration = false;
                                        break;
                                }
                            }
                        } while (retryIteration);
                    }
             

            


        // Create an Excel file to store data
        Console.WriteLine("Creating an Excel workbook to store the extracted data...");
                var workbook = new XLWorkbook();
                var worksheet = workbook.AddWorksheet("Device Data");

                // Set headers for the Excel file
                Console.WriteLine("Setting headers for the Excel file...");
                worksheet.Cell(1, 1).Value = "UDI-DI";
                //Manufacturer details 
                worksheet.Cell(1, 2).Value = "Version";
                worksheet.Cell(1, 3).Value = "Last Update Date";
                worksheet.Cell(1, 4).Value = "Actor/Organisation name";
                worksheet.Cell(1, 5).Value = "Actor ID/SRN";
                worksheet.Cell(1, 6).Value = "Address";
                worksheet.Cell(1, 7).Value = "Country";
                worksheet.Cell(1, 8).Value = "Telephone number";
                worksheet.Cell(1, 9).Value = "Email";
                //Basic UDI-DI details
                worksheet.Cell(1, 10).Value = "Version";
                worksheet.Cell(1, 11).Value = "Last update date";
                worksheet.Cell(1, 12).Value = "Applicable legislation";
                worksheet.Cell(1, 13).Value = "Basic UDI-DI/EUDAMED DI / Issuing entity";
                worksheet.Cell(1, 14).Value = "Kit";
                worksheet.Cell(1, 15).Value = "System/Procedure which is a device in itself";
                worksheet.Cell(1, 16).Value = "Authorised representative";
                worksheet.Cell(1, 17).Value = "Special device type";
                worksheet.Cell(1, 18).Value = "Risk class";
                worksheet.Cell(1, 19).Value = "Implantable";
                worksheet.Cell(1, 20).Value = "Is the device a suture, staple, dental filling, dental brace, tooth crown, screw, wedge, plate, wire, pin, clip or connector?";
                worksheet.Cell(1, 21).Value = "Measuring function";
                worksheet.Cell(1, 22).Value = "Reusable surgical instrument";
                worksheet.Cell(1, 23).Value = "Active device";
                worksheet.Cell(1, 24).Value = "Device intended to administer and / or remove medicinal product";
                worksheet.Cell(1, 25).Value = "Companion diagnostic";
                worksheet.Cell(1, 26).Value = "Near patient testing";
                worksheet.Cell(1, 27).Value = "Patient self testing";
                worksheet.Cell(1, 28).Value = "Professional testing";
                worksheet.Cell(1, 29).Value = "Reagent";
                worksheet.Cell(1, 30).Value = "Instrument";
                worksheet.Cell(1, 31).Value = "Device model";
                worksheet.Cell(1, 32).Value = "Device name";
                //Tissues and cells
                worksheet.Cell(1, 33).Value = "Presence of human tissues and cells or their derivatives";
                worksheet.Cell(1, 34).Value = "Presence of animal tissues and cells or their derivatives";
                worksheet.Cell(1, 35).Value = "Presence of cells or substance of microbial origin";
                //Information on substances
                worksheet.Cell(1, 36).Value = "Presence of a substance which, if used separately, may be considered to be a medicinal product";
                worksheet.Cell(1, 37).Value = "Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma";
                //UDI - DI details
                worksheet.Cell(1, 38).Value = "Version";
                worksheet.Cell(1, 39).Value = "Last update date";

                worksheet.Cell(1, 40).Value = "UDI-DI code / Issuing entity";
                worksheet.Cell(1, 41).Value = "Status";
                worksheet.Cell(1, 42).Value = "UDI-DI from another entity (secondary)";
                worksheet.Cell(1, 43).Value = "Nomenclature code(s)";
                worksheet.Cell(1, 44).Value = "Name/Trade name(s)";
                worksheet.Cell(1, 45).Value = "Reference / Catalogue number";
                worksheet.Cell(1, 46).Value = "Direct marking DI";
                worksheet.Cell(1, 47).Value = "Unit of Use DI";
                worksheet.Cell(1, 48).Value = "Quantity of device";
                worksheet.Cell(1, 49).Value = "Type of UDI-PI";
                worksheet.Cell(1, 50).Value = "Additional Product description";
                worksheet.Cell(1, 51).Value = "Additional information url";
                worksheet.Cell(1, 52).Value = "Clinical sizes";
                worksheet.Cell(1, 53).Value = "Labelled as single use";
                worksheet.Cell(1, 54).Value = "Maximum number of reuses";
                worksheet.Cell(1, 55).Value = "Need for sterilisation before use";
                worksheet.Cell(1, 56).Value = "Device labelled as sterile";
                worksheet.Cell(1, 57).Value = "Containing Latex";
                worksheet.Cell(1, 58).Value = "Storage and handling conditions";
                worksheet.Cell(1, 59).Value = "Critical warnings or contra-indications";
                worksheet.Cell(1, 60).Value = "Reprocessesed single use device";
                worksheet.Cell(1, 61).Value = "Intended purpose other than medical (Annex XVI)";
                worksheet.Cell(1, 62).Value = "Member state of the placing on the EU market of the device";
                worksheet.Cell(1, 63).Value = "Presence of a substance which, if used separately, may be considered to be a medicinal product";
                //Market distribution 
                worksheet.Cell(1, 64).Value = "Version ";
                worksheet.Cell(1, 65).Value = "Last update date";
                worksheet.Cell(1, 66).Value = "Member State where the device is or is to be made available";
                //SS(C)P
                worksheet.Cell(1, 67).Value = "SS(C)P Reference number";
                worksheet.Cell(1, 68).Value = "SS(C)P revision number";
                worksheet.Cell(1, 69).Value = "Issue date";
                //Certificate
                worksheet.Cell(1, 70).Value = "Certificates numbers";

                //int rowNum = 2;

                // Start iterating over the rows of the table
                Console.WriteLine("Starting to iterate over the table rows...");
                int excelRowIndex = 2;


                for (int currentPage = 2208; currentPage <= totalPages; currentPage++)
                {

                    var tableRows = driver.FindElements(By.CssSelector("table tbody tr"));
                    for (int i = 0; i < tableRows.Count; i++)
                        do
                        {
                            try
                            {
                                retryIteration = false;
                                // Refresh the list of rows on each iteration
                                tableRows = driver.FindElements(By.CssSelector("table tbody tr"));

                                Console.WriteLine($"Clicking the 'View detail' button for website row {i + 1}, saving to Excel row {excelRowIndex}...");
                                var viewDetailButton = tableRows[i].FindElement(By.XPath(".//button[@title='View detail']"));


                                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", viewDetailButton);
                                viewDetailButton.Click();


                                // Wait for the detail page to load
                                // Console.WriteLine("Waiting for the detail page to load...");
                                // wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("div.ecl-container")));
                                // Console.WritseLine("Div with class 'ecl-container' has loaded.");

                                // WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                                var accordionElements = wait.Until(d => d.FindElements(By.XPath("//div[@class='mb-5']")));
                                Console.WriteLine("Details has loaded.");
                                //
                                // Extract the UDI-DI text


                                //var udiElement = wait.Until(d => d.FindElement(By.XPath("//h1[contains(text(), 'UDI-DI')]"))).Text;
                                //var udiText = udiElement.Split(':').Last().Trim();
                                //Console.WriteLine("UDI-DI: " + udiText);




                                //
                                // Extract the Version
                                //

                                var versionElement = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("(//ul[@id='versionStatus']/li/strong)[1]")));
                                var versionText = versionElement.Text;
                                Console.WriteLine("Version: " + versionText);

                                //string versionXpath = "(//ul[@id='versionStatus']/li/strong)[1]";
                                //string versionText = driver.FindElement(By.XPath(versionXpath)).Text;
                                //Console.WriteLine("Version: " + versionText);


                                // Extract the Last Update Date
                                var lastUpdateElement = wait.Until(d => d.FindElement(By.XPath("(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[1]")));
                                var lastUpdateText = lastUpdateElement.Text.Replace("Last update date: ", "").Trim();
                                Console.WriteLine("Last Update Date: " + lastUpdateText);

                                // Extract the Actor/Organisation name
                                var actorNameElement = wait.Until(d => d.FindElement(By.XPath("(//dt[text()='Actor/Organisation name']/following-sibling::dd/div)[1]")));
                                var actorNameText = actorNameElement.Text;
                                Console.WriteLine("Actor/Organisation Name: " + actorNameText);

                                // Extract the Actor ID/SRN
                                var actorIdElement = wait.Until(d => d.FindElement(By.XPath("//dt[text()='Actor ID/SRN']/following-sibling::dd/div")));
                                var actorIdText = actorIdElement.Text.Trim();
                                Console.WriteLine("Actor ID/SRN: " + actorIdText);

                                // Extract the Address
                                var addressElement = wait.Until(d => d.FindElement(By.XPath("//dt[text()='Address']/following-sibling::dd/div")));
                                var addressText = addressElement.Text.Trim();
                                Console.WriteLine("Address: " + addressText);

                                // Extract the Country
                                var countryElement = wait.Until(d => d.FindElement(By.XPath("//dt[text()='Country']/following-sibling::dd/div")));
                                var countryText = countryElement.Text.Trim();
                                Console.WriteLine("Country: " + countryText);

                                // Extract the Telephone number
                                var telephoneElement = wait.Until(d => d.FindElement(By.XPath("//dt[text()='Telephone number']/following-sibling::dd/div")));
                                var telephoneText = telephoneElement.Text.Trim();
                                Console.WriteLine("Telephone Number: " + telephoneText);

                                // Extract the Email
                                var emailElement = wait.Until(d => d.FindElement(By.XPath("//dt[text()='Email']/following-sibling::dd/div")));
                                var emailText = emailElement.Text.Trim();
                                Console.WriteLine("Email: " + emailText);
                                //
                                ////Basic UDI-DI details
                                //
                                var wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                                //
                                // Extract Version

                                string versionXpath2 = "(//ul[@id='versionStatus']/li/strong)[2]";
                                string versionText2 = "";

                                try
                                {
                                    versionText2 = driver.FindElement(By.XPath(versionXpath2)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Version 2 not found. Leaving it empty.");
                                }

                                Console.WriteLine("Version: " + versionText2);

                                // Extract Last Update Date
                                var lastUpdateElement2 = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[2]")));
                                var lastUpdateText2 = lastUpdateElement.Text.Replace("Last update date: ", "").Trim();
                                Console.WriteLine("Last Update Date: " + lastUpdateText2);

                                // Extract Applicable Legislation

                                var legislationElement = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Applicable legislation')]/following-sibling::dd/div")));
                                var applicableLegislation = legislationElement.Text;
                                Console.WriteLine("Applicable Legislation: " + applicableLegislation);

                                // Extract Basic UDI-DI/EUDAMED DI / Issuing Entity

                                //var udiElement_basic = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//dt[contains(text(), 'Basic UDI-DI/EUDAMED DI / Issuing entity')]/following-sibling::dd/div")));
                                //var udiText_basic = udiElement_basic.Text.Trim();
                                //Console.WriteLine("Basic UDI-DI/EUDAMED DI / Issuing Entity: " + udiText_basic);

                                string udiElement_basic = "//dt[contains(text(), 'UDI-DI/EUDAMED')]/following-sibling::dd/div";
                                string udiText_basic = "";

                                try
                                {
                                    udiText_basic = driver.FindElement(By.XPath(udiElement_basic)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Basic UDI-DI/EUDAMED DI / Issuing Entity not found. Leaving it empty.");
                                }

                                Console.WriteLine("Basic UDI-DI/EUDAMED DI / Issuing Entity: " + udiText_basic);

                                ////Kit

                                string kitElement = "//dt[contains(text(), 'Kit')]/following-sibling::dd/div";
                                string kitText = "";

                                try
                                {
                                    kitText = driver.FindElement(By.XPath(kitElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Kit not found. Leaving it empty.");
                                }

                                Console.WriteLine("Kit: " + kitText);

                                //// Extract System/Procedure
                                //
                                string systemProcedureElement = "//dt[contains(text(), 'System')]/following-sibling::dd/div";
                                string systemProcedure = "";

                                try
                                {
                                    systemProcedure = driver.FindElement(By.XPath(systemProcedureElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("System/Procedure which is a device in itself not found. Leaving it empty.");
                                }
                                Console.WriteLine("System/Procedure which is a device in itself: " + systemProcedure);
                                //
                                //// Extract Authorised Representative
                                //
                                string authorisedRepElement = "//dt[contains(text(), 'Authorised representative')]/following-sibling::dd/div";
                                string authorisedRep = "";

                                try
                                {
                                    authorisedRep = driver.FindElement(By.XPath(authorisedRepElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Authorised representative not found. Leaving it empty.");
                                }

                                Console.WriteLine("Authorised representative: " + authorisedRep);


                                ////Special device type

                                string specDevTypeElement = "//dt[contains(text(), 'Special device Type')]/following-sibling::dd/div";
                                string specDevTypeText = "";

                                try
                                {
                                    specDevTypeText = driver.FindElement(By.XPath(specDevTypeElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Special device type not found. Leaving it empty.");
                                }

                                Console.WriteLine("Special device type: " + specDevTypeText);



                                //// Extract Risk Class
                                string riskClassElement = "//dt[contains(text(), 'Risk class')]/following-sibling::dd/div";
                                string riskClass = "";
                                try
                                {
                                    riskClass = driver.FindElement(By.XPath(riskClassElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Implantable not found. Leaving it empty.");
                                }

                                Console.WriteLine("Risk Class: " + riskClass);

                                //// Extract Implantable

                                string implantableElement = "//dt[contains(text(), 'Implantable')]/following-sibling::dd/div";
                                string implantable = "";

                                try
                                {
                                    implantable = driver.FindElement(By.XPath(implantableElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Implantable not found. Leaving it empty.");
                                }

                                Console.WriteLine("Implantable: " + implantable);


                                //// Extract Suture/Staple Device

                                string sutureElement = "//dt[contains(text(), 'Is the device a suture, ')]/following-sibling::dd/div";
                                string sutureDevice = "";

                                try
                                {
                                    sutureDevice = driver.FindElement(By.XPath(sutureElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Suture device status not found. Leaving it empty.");
                                }
                                Console.WriteLine("Is the device a suture/staple/etc: " + sutureDevice);
                                //
                                //// Extract Measuring Function

                                string measuringFunctionElement = "//dt[contains(text(), 'Measuring function')]/following-sibling::dd/div";
                                string measuringFunction = "";

                                try
                                {
                                    measuringFunction = driver.FindElement(By.XPath(measuringFunctionElement)).Text.Trim();
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Measuring Function not found. Leaving it empty.");
                                }
                                Console.WriteLine("Measuring Function: " + measuringFunction);
                                //
                                //// Extract Reusable Surgical Instrument

                                string reusableInstrumentElement = "//dt[contains(text(), 'Reusable surgical instrument')]/following-sibling::dd/div";
                                string reusableInstrument = "";

                                try
                                {
                                    reusableInstrument = driver.FindElement(By.XPath(reusableInstrumentElement)).Text.Trim();
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Reusable Surgical Instrument not found. Leaving it empty.");
                                }

                                Console.WriteLine("Reusable Surgical Instrument: " + reusableInstrument);
                                //
                                // Extract Active Device

                                string activeDeviceElement = "//dt[contains(text(), 'Active device')]/following-sibling::dd/div";
                                string activeDevice = "";

                                try
                                {
                                    activeDevice = driver.FindElement(By.XPath(activeDeviceElement)).Text.Trim();
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Active Device not found. Leaving it empty.");
                                }
                                Console.WriteLine("Active Device: " + activeDevice);

                                // Extract Device Intended to Administer Medicinal Product

                                string adminDeviceElement = "//dt[contains(text(), 'Device intended to administer and / or remove medicinal product')]/following-sibling::dd/div";
                                string adminDevice = "";

                                try
                                {
                                    adminDevice = driver.FindElement(By.XPath(adminDeviceElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Device Intended to Administer Medicinal Product not found. Leaving it empty.");
                                }
                                Console.WriteLine("Device Intended to Administer Medicinal Product: " + adminDevice);



                                ////Companion diagnostic

                                string compDiagElement = "//dt[contains(text(), 'Companion diagnostic')]/following-sibling::dd/div";
                                string compDiagText = "";

                                try
                                {
                                    compDiagText = driver.FindElement(By.XPath(compDiagElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Companion diagnostic not found. Leaving it empty.");
                                }

                                Console.WriteLine("Companion diagnostic: " + compDiagText);

                                ////Near patient testing

                                string nearPatTestElement = "//dt[contains(text(), 'Near patient testing')]/following-sibling::dd/div";
                                string nearPatTestText = "";

                                try
                                {
                                    nearPatTestText = driver.FindElement(By.XPath(nearPatTestElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Near patient testing not found. Leaving it empty.");
                                }

                                Console.WriteLine("Near patient testing: " + nearPatTestText);


                                ////Patient self testing
                                ///


                                string patSelfTestElement = "//dt[contains(text(), 'Patient self testing')]/following-sibling::dd/div";
                                string patSelfTestText = "";

                                try
                                {
                                    patSelfTestText = driver.FindElement(By.XPath(patSelfTestElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Near patient testing not found. Leaving it empty.");
                                }

                                Console.WriteLine("Patient self testing: " + patSelfTestText);


                                ////Professional testing

                                string profTestElement = "//dt[contains(text(), 'Professional testing')]/following-sibling::dd/div";
                                string profTestText = "";

                                try
                                {
                                    profTestText = driver.FindElement(By.XPath(profTestElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Professional testing not found. Leaving it empty.");
                                }

                                Console.WriteLine("Professional testing: " + profTestText);


                                ////Reagent

                                string reagentElement = "//dt[contains(text(), 'Reagent')]/following-sibling::dd/div";
                                string reagentText = "";

                                try
                                {
                                    reagentText = driver.FindElement(By.XPath(reagentElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Reagent not found. Leaving it empty.");
                                }

                                Console.WriteLine("Reagent: " + reagentText);


                                ////Instrument



                                string InstrumentElement = "//dt[contains(text(), 'Instrument')]/following-sibling::dd/div";
                                string InstrumentText = "";

                                try
                                {
                                    InstrumentText = driver.FindElement(By.XPath(InstrumentElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Instrument not found. Leaving it empty.");
                                }

                                Console.WriteLine("Instrument: " + InstrumentText);


                                ////Device model 

                                string deviceModelElement = "//dt[contains(text(), 'Device model')]/following-sibling::dd/div";
                                string deviceModelText = "";

                                try
                                {
                                    deviceModelText = driver.FindElement(By.XPath(deviceModelElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Device model not found. Leaving it empty.");
                                }

                                Console.WriteLine("Device model: " + deviceModelText);







                                // Extract Device Name
                                var deviceNameElement = wait.Until(d => d.FindElement(By.XPath("//dt[contains(text(), 'Device name')]/following-sibling::dd/div")));
                                var deviceName = deviceNameElement.Text.Trim();
                                Console.WriteLine("Device Name: " + deviceName);

                                ////Tissues and cells
                                //// Extract "Presence of human tissues and cells or their derivatives"
                                string humanTissuesXpath = "//dt[text()='Presence of human tissues and cells or their derivatives']/following-sibling::dd/div";
                                string presenceOfHumanTissues = driver.FindElement(By.XPath(humanTissuesXpath)).Text;
                                Console.WriteLine("Presence of human tissues and cells or their derivatives: " + presenceOfHumanTissues);

                                // Extract the "Presence of animal tissues and cells or their derivatives"
                                string animalTissuesXpath = "//dt[text()='Presence of animal tissues and cells or their derivatives']/following-sibling::dd/div";
                                string presenceOfAnimalTissues = driver.FindElement(By.XPath(animalTissuesXpath)).Text;
                                Console.WriteLine("Presence of animal tissues and cells or their derivatives: " + presenceOfAnimalTissues);


                                ////Presence of cells or substances of microbial origin

                                string microbialElement = "//dt[contains(text(), 'Presence of cells or substances of microbial origin')]/following-sibling::dd/div";
                                string microbialText = "";

                                try
                                {
                                    microbialText = driver.FindElement(By.XPath(microbialElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Presence of cells or substances of microbial origin not found. Leaving it empty.");
                                }

                                Console.WriteLine("Presence of cells or substances of microbial origin: " + microbialText);

                                //Information on substances

                                // Extract the "Presence of a substance which, if used separately, may be considered to be a medicinal product"
                                string medicinalProductXpath = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product']/following-sibling::dd/div";
                                string presenceOfMedicinalProduct = driver.FindElement(By.XPath(medicinalProductXpath)).Text;
                                Console.WriteLine("Presence of a substance which, if used separately, may be considered to be a medicinal product: " + presenceOfMedicinalProduct);

                                // Extract the "Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma"
                                string bloodPlasmaProductXpath = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma']/following-sibling::dd/div";
                                string presenceOfBloodPlasmaProduct = driver.FindElement(By.XPath(bloodPlasmaProductXpath)).Text;
                                Console.WriteLine("Presence of a substance which, if used separately, may be considered to be a medicinal product derived from human blood or human plasma: " + presenceOfBloodPlasmaProduct);

                                ////UDI - DI details
                                //
                                // Extract the "Version 1 (Current)"
                                string versionXpath3 = "(//ul[@id='versionStatus']/li/strong)[3]";
                                string versionText3 = "";

                                try
                                {
                                    versionText3 = driver.FindElement(By.XPath(versionXpath3)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Version 3 not found. Leaving it empty.");
                                }

                                Console.WriteLine("Version: " + versionText3);

                                //// Extract the "Last update date"
                                //string lastUpdateXpath = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[3]";
                                //string lastUpdateText3 = driver.FindElement(By.XPath(lastUpdateXpath)).Text;
                                //Console.WriteLine("Last update date: " + lastUpdateText3);
                                string lastUpdateXpath3 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[3]";
                                string lastUpdateText3 = "";

                                try
                                {
                                    lastUpdateText3 = driver.FindElement(By.XPath(lastUpdateXpath3)).Text.Replace("Last update date: ", "").Trim(); ;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Last Update Date 3 not found. Leaving it empty.");
                                }

                                Console.WriteLine("Last Update Date: " + lastUpdateText3);

                                //
                                //// Extract the "UDI-DI code / Issuing entity"
                                string udiDiXpath = "//dt[text()='UDI-DI code / Issuing entity']/following-sibling::dd/div";
                                string udiDi = driver.FindElement(By.XPath(udiDiXpath)).Text;
                                Console.WriteLine("UDI-DI code / Issuing entity: " + udiDi);

                                //// Extract the "Status"
                                string statusXpath = "//dt[text()='Status']/following-sibling::dd/div";
                                string status = driver.FindElement(By.XPath(statusXpath)).Text;
                                Console.WriteLine("Status: " + status);

                                //// Extract the "UDI-DI from another entity (secondary)"
                                string secondaryUdiXpath = "//dt[text()='UDI-DI from another entity (secondary)']/following-sibling::dd/div";
                                string secondaryUdi = "";

                                try
                                {
                                    secondaryUdi = driver.FindElement(By.XPath(secondaryUdiXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("UDI-DI from another entity (secondary) not found. Leaving it empty.");
                                }

                                Console.WriteLine("UDI-DI from another entity (secondary): " + secondaryUdi);

                                //// Extract the "Nomenclature code(s)"
                                string nomenclatureCodeXpath = "//dt[text()='Nomenclature code(s)']/following-sibling::dd/div";
                                string nomenclatureCode = driver.FindElement(By.XPath(nomenclatureCodeXpath)).Text;
                                Console.WriteLine("Nomenclature code(s): " + nomenclatureCode);

                                //// Extract the "Name/Trade name(s)"
                                string tradeNameXpath = "//dt[text()='Name/Trade name(s)']/following-sibling::dd/div";
                                string tradeName = driver.FindElement(By.XPath(tradeNameXpath)).Text;
                                Console.WriteLine("Name/Trade name(s): " + tradeName);

                                //// Extract the "Reference / Catalogue number"
                                string catalogueNumberXpath = "//dt[text()='Reference / Catalogue number']/following-sibling::dd/div";
                                string catalogueNumber = driver.FindElement(By.XPath(catalogueNumberXpath)).Text;
                                Console.WriteLine("Reference / Catalogue number: " + catalogueNumber);

                                // Extract the "Direct marking DI"
                                string directMarkingXpath = "//dt[text()='Direct marking DI']/following-sibling::dd/div";
                                string directMarking = "";


                                try
                                {
                                    directMarking = driver.FindElement(By.XPath(directMarkingXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Direct marking DI not found. Leaving it empty.");
                                }
                                Console.WriteLine("Direct marking DI: " + directMarking);

                                ////Unit of use

                                string unitOfUseElement = "//dt[contains(text(), 'Unit of use')]/following-sibling::dd/div";
                                string unitOfUseText = "";

                                try
                                {
                                    unitOfUseText = driver.FindElement(By.XPath(unitOfUseElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Unit of Use not found. Leaving it empty.");
                                }

                                Console.WriteLine("Unit of Use: " + unitOfUseText);


                                // Extract the "Quantity of device"
                                string quantityXpath = "//dt[text()='Quantity of device']/following-sibling::dd/div";
                                string quantity = "";

                                try
                                {
                                    quantity = driver.FindElement(By.XPath(quantityXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Quantity of device not found. Leaving it empty.");
                                }

                                Console.WriteLine("Quantity of device: " + quantity);
                                //
                                //// Extract the "Type of UDI-PI"
                                string udiPiXpath = "//dt[text()='Type of UDI-PI']/following-sibling::dd/div";
                                string udiPi = "";

                                try
                                {
                                    udiPi = driver.FindElement(By.XPath(udiPiXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Type of UDI-PI not found. Leaving it empty.");
                                }
                                Console.WriteLine("Type of UDI-PI: " + udiPi);
                                //
                                //// Extract the "Additional Product description"
                                string additionalDescriptionXpath = "//dt[text()='Additional Product description']/following-sibling::dd/div";
                                string additionalDescription = "";

                                try
                                {
                                    additionalDescription = driver.FindElement(By.XPath(additionalDescriptionXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Additional Product description not found. Leaving it empty.");
                                }
                                Console.WriteLine("Additional Product description: " + additionalDescription);
                                //
                                //// Extract the "Additional information url"
                                string infoUrlXpath = "//dt[text()='Additional information url']/following-sibling::dd/div";
                                string infoUrl = "";

                                try
                                {
                                    infoUrl = driver.FindElement(By.XPath(infoUrlXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Additional information url not found. Leaving it empty.");
                                }
                                Console.WriteLine("Additional information url: " + infoUrl);
                                //
                                //// Extract the "Clinical sizes"
                                string clinicalSizesXpath = "//dt[text()='Clinical sizes']/following-sibling::dd/div";
                                string clinicalSizes = "";

                                try
                                {
                                    clinicalSizes = driver.FindElement(By.XPath(clinicalSizesXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Clinical sizes not found. Leaving it empty.");
                                }
                                Console.WriteLine("Clinical sizes: " + clinicalSizes);

                                // Extract the "Labelled as single use"
                                string singleUseXpath = "//dt[text()='Labelled as single use']/following-sibling::dd/div";
                                string singleUse = "";

                                try
                                {
                                    singleUse = driver.FindElement(By.XPath(singleUseXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Labelled as single use not found. Leaving it empty.");
                                }
                                Console.WriteLine("Labelled as single use: " + singleUse);


                                // Extract the "Maximum number of reuses"
                                string maxNoReusesElement = "//dt[text()='Maximum number of reuses']/following-sibling::dd/div";
                                string maxNoReusesText = "";

                                try
                                {
                                    maxNoReusesText = driver.FindElement(By.XPath(maxNoReusesElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Maximum number of reuses not found. Leaving it empty.");
                                }
                                Console.WriteLine("Maximum number of reuses: " + maxNoReusesText);






                                // Extract the "Need for sterilisation before use"
                                string sterilisationXpath = "//dt[text()='Need for sterilisation before use']/following-sibling::dd/div";
                                string sterilisation = "";

                                try
                                {
                                    sterilisation = driver.FindElement(By.XPath(sterilisationXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Need for sterilisation before use not found. Leaving it empty.");
                                }
                                Console.WriteLine("Need for sterilisation before use: " + sterilisation);

                                // Extract the "Device labelled as sterile"
                                string sterileXpath = "//dt[text()='Device labelled as sterile']/following-sibling::dd/div";
                                string sterile = "";

                                try
                                {
                                    sterile = driver.FindElement(By.XPath(sterileXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Device labelled as sterile not found. Leaving it empty.");
                                }
                                Console.WriteLine("Device labelled as sterile: " + sterile);

                                // Extract the "Containing Latex"
                                string latexXpath = "//dt[text()='Containing Latex']/following-sibling::dd/div";
                                string latex = "";

                                try
                                {
                                    latex = driver.FindElement(By.XPath(latexXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Containing Latex not found. Leaving it empty.");
                                }
                                Console.WriteLine("Containing Latex: " + latex);


                                // Extract the "Storage and handling conditions "
                                string handlingCondElement = "//dt[text()='Storage and handling conditions']/following-sibling::dd/div";
                                string handlingCondText = "";

                                try
                                {
                                    handlingCondText = driver.FindElement(By.XPath(handlingCondElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Storage and handling conditions not found. Leaving it empty.");
                                }
                                Console.WriteLine("Storage and handling conditions: " + handlingCondText);

                                // Extract the "Critical warnings or contra-indications"
                                string warningsXpath = "//dt[text()='Critical warnings or contra-indications']/following-sibling::dd/div";
                                string warnings = "";

                                try
                                {
                                    warnings = driver.FindElement(By.XPath(warningsXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Critical warnings or contra-indications not found. Leaving it empty.");
                                }
                                Console.WriteLine("Critical warnings or contra-indications: " + warnings);

                                // Extract the "Do not re-use"
                                string doNotReuseXpath = "//dt[text()='Critical warnings or contra-indications']/following-sibling::dd//li[text()='Do not re-use']";
                                string doNotReuse = "";

                                try
                                {
                                    doNotReuse = driver.FindElement(By.XPath(doNotReuseXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Do not re-use not found. Leaving it empty.");
                                }
                                Console.WriteLine("Do not re-use: " + doNotReuse);

                                // Extract the "Reprocessed single use device"
                                string reprocessedXpath = "//dt[contains(text(), 'Reprocessesed single use device')]/following-sibling::dd/div";
                                string reprocessed = "";

                                try
                                {
                                    reprocessed = driver.FindElement(By.XPath(reprocessedXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Reprocessed single use device not found. Leaving it empty.");
                                }
                                Console.WriteLine("Reprocessed single use device: " + reprocessed);

                                // Extract the "Intended purpose other than medical (Annex XVI)"
                                string intendedPurposeXpath = "//dt[contains(text(), 'Intended purpose other than medical')]/following-sibling::dd/div";
                                string intendedPurpose = "";

                                try
                                {
                                    intendedPurpose = driver.FindElement(By.XPath(intendedPurposeXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Intended purpose other than medical (Annex XVI) not found. Leaving it empty.");
                                }
                                Console.WriteLine("Intended purpose other than medical (Annex XVI): " + intendedPurpose);

                                // Extract the "Member state of the placing on the EU market of the device"
                                string memberStateXpath = "//dt[text()='Member state of the placing on the EU market of the device']/following-sibling::dd/div";
                                string memberState = "";

                                try
                                {
                                    memberState = driver.FindElement(By.XPath(memberStateXpath)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Member state of the placing on the EU market of the device not found. Leaving it empty.");
                                }
                                Console.WriteLine("Member state of the placing on the EU market of the device: " + memberState);

                                // Extract the "Presence of a substance which, if used separately, may be considered to be a medicinal product"
                                string medProdElement = "//dt[text()='Presence of a substance which, if used separately, may be considered to be a medicinal product']/following-sibling::dd/div";
                                string medProdText = "";

                                try
                                {
                                    medProdText = driver.FindElement(By.XPath(medProdElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Presence of a substance which, if used separately, may be considered to be a medicinal product not found. Leaving it empty.");
                                }
                                Console.WriteLine("Presence of a substance which, if used separately, may be considered to be a medicinal product: " + medProdText);
                                //
                                //// Market distribution
                                //
                                // Extract the "Version 1 (Current)"
                                string versionXpath4 = "(//ul[@id='versionStatus']/li/strong)[4]";
                                string versionText4 = "";

                                try
                                {
                                    versionText4 = driver.FindElement(By.XPath(versionXpath4)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Version 4 not found. Leaving it empty.");
                                }

                                Console.WriteLine("Version: " + versionText4);

                                //
                                //// Extract the "Last update date"
                                //

                                string lastUpdateXpath4 = "(//ul[@id='versionStatus']/li[contains(text(), 'Last update date:')])[4]";
                                string lastUpdateText4 = "";

                                try
                                {
                                    lastUpdateText4 = driver.FindElement(By.XPath(lastUpdateXpath4)).Text.Replace("Last update date: ", "").Trim(); ;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Last Update Date 4 not found. Leaving it empty.");
                                }

                                Console.WriteLine("Last Update Date: " + lastUpdateText4);
                                //
                                //// Extract the "Member State where the device is or is to be made available"
                                //string memberStateXpath2 = "//dt[text()='Member State where the device is or is to be made available']/following-sibling::dd//ul";
                                //string memberStateAvailab = driver.FindElement(By.XPath(memberStateXpath2)).Text;
                                //Console.WriteLine("Member State where the device is or is to be made available: " + memberStateAvailab);

                                string memberStateXpath2 = "//dt[text()='Member State where the device is or is to be made available']/following-sibling::dd//ul";
                                string memberStateAvailab = "";

                                try
                                {
                                    memberStateAvailab = driver.FindElement(By.XPath(memberStateXpath2)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("Member State not found. Leaving it empty.");
                                }

                                Console.WriteLine("Member State: " + memberStateAvailab);


                                //
                                //// Extract the "SS(C)P Reference number"
                                //

                                string refNoElement = "//dt[text()='SS(C)P Reference number']/following-sibling::dd/div";
                                string refNoText = "";

                                try
                                {
                                    refNoText = driver.FindElement(By.XPath(refNoElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("SS(C)P Reference number not found. Leaving it empty.");
                                }
                                Console.WriteLine("SS(C)P Reference number: " + refNoText);


                                //
                                //// Extract the "SS(C)P revision number"
                                //

                                string revNoElement = "//dt[text()='SS(C)P revision number']/following-sibling::dd/div";
                                string revNoText = "";

                                try
                                {
                                    revNoText = driver.FindElement(By.XPath(revNoElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("SS(C)P revision number not found. Leaving it empty.");
                                }
                                Console.WriteLine("SS(C)P revision number: " + revNoText);


                                //
                                //// Extract the "Issue date"
                                //

                                string issueDateElement = "//dt[text()='Issue date']/following-sibling::dd/div";
                                string issueDateText = "";

                                try
                                {
                                    issueDateText = driver.FindElement(By.XPath(issueDateElement)).Text;
                                }
                                catch (NoSuchElementException)
                                {
                                    // If the element is not found, leave versionText4 as empty
                                    Console.WriteLine("SS(C)P issue date not found. Leaving it empty.");
                                }
                                Console.WriteLine("SS(C)P issue date: " + issueDateText);

                                //
                                //// Extract the "Certificates numbers"
                                //


                                // XPath for the certificate headers
                                string certificateNoElement = "//h2[text()='Certificates']/following-sibling::div[1]//mat-expansion-panel-header";
                                string certificateNoElement2 = "//h2[text()='Certificates']/following-sibling::div[1]//mat-expansion-panel/div/div/div";
                                string certificateNoText = "";

                                try
                                {
                                    // Find all matching elements
                                    var certificateElements = driver.FindElements(By.XPath(certificateNoElement));
                                    var certificateElements2 = driver.FindElements(By.XPath(certificateNoElement2));

                                    if (certificateElements.Count > 0 || certificateElements2.Count > 0)
                                    {
                                        // Extract text from both sets of elements and concatenate them with " % "
                                        var certificateTexts = certificateElements.Select(el => el.Text)
                                                                .Concat(certificateElements2.Select(el => el.Text));

                                        certificateNoText = string.Join("  %  ", certificateTexts) + " % ";
                                    }
                                    else
                                    {
                                        Console.WriteLine("Certificates numbers not found. Leaving it empty.");
                                    }
                                }
                                catch (NoSuchElementException)
                                {
                                    Console.WriteLine("Certificates numbers not found. Leaving it empty.");
                                }


                                // Print the final output
                                Console.WriteLine("Certificates numbers: " + certificateNoText);



                                //// Save extracted data to Excel
                                Console.WriteLine($"Saving data for UDI-DI: {udiDi}...");

                                worksheet.Cell(excelRowIndex, 2).Value = versionText;
                                worksheet.Cell(excelRowIndex, 3).Value = lastUpdateText;
                                worksheet.Cell(excelRowIndex, 4).Value = actorNameText;
                                worksheet.Cell(excelRowIndex, 5).Value = actorIdText;
                                worksheet.Cell(excelRowIndex, 6).Value = addressText;
                                worksheet.Cell(excelRowIndex, 7).Value = countryText;
                                worksheet.Cell(excelRowIndex, 8).Value = telephoneText;
                                worksheet.Cell(excelRowIndex, 9).Value = emailText;
                                //
                                ////Basic UDI-DI
                                //
                                worksheet.Cell(excelRowIndex, 10).Value = versionText2;
                                worksheet.Cell(excelRowIndex, 11).Value = lastUpdateText2;
                                worksheet.Cell(excelRowIndex, 12).Value = applicableLegislation;
                                worksheet.Cell(excelRowIndex, 13).Value = udiText_basic;
                                worksheet.Cell(excelRowIndex, 14).Value = kitText;
                                worksheet.Cell(excelRowIndex, 15).Value = systemProcedure;
                                worksheet.Cell(excelRowIndex, 16).Value = authorisedRep;
                                worksheet.Cell(excelRowIndex, 17).Value = specDevTypeText;
                                worksheet.Cell(excelRowIndex, 18).Value = riskClass;
                                worksheet.Cell(excelRowIndex, 19).Value = implantable;
                                worksheet.Cell(excelRowIndex, 20).Value = sutureDevice;
                                worksheet.Cell(excelRowIndex, 21).Value = measuringFunction;
                                worksheet.Cell(excelRowIndex, 22).Value = reusableInstrument;
                                worksheet.Cell(excelRowIndex, 23).Value = activeDevice;
                                worksheet.Cell(excelRowIndex, 24).Value = adminDevice;
                                worksheet.Cell(excelRowIndex, 25).Value = compDiagText;
                                worksheet.Cell(excelRowIndex, 26).Value = nearPatTestText;
                                worksheet.Cell(excelRowIndex, 27).Value = patSelfTestText;
                                worksheet.Cell(excelRowIndex, 28).Value = profTestText;
                                worksheet.Cell(excelRowIndex, 29).Value = reagentText;
                                worksheet.Cell(excelRowIndex, 30).Value = InstrumentText;
                                worksheet.Cell(excelRowIndex, 31).Value = deviceModelText;
                                worksheet.Cell(excelRowIndex, 32).Value = deviceName;
                                //
                                ////Tissues and cells

                                worksheet.Cell(excelRowIndex, 33).Value = presenceOfHumanTissues;
                                worksheet.Cell(excelRowIndex, 34).Value = presenceOfAnimalTissues;
                                worksheet.Cell(excelRowIndex, 35).Value = microbialText;
                                //
                                ////Information on Substances

                                worksheet.Cell(excelRowIndex, 36).Value = presenceOfMedicinalProduct;
                                worksheet.Cell(excelRowIndex, 37).Value = presenceOfBloodPlasmaProduct;
                                //
                                ////UDI-DI details
                                //
                                worksheet.Cell(excelRowIndex, 38).Value = versionText3;
                                worksheet.Cell(excelRowIndex, 39).Value = lastUpdateText3;
                                worksheet.Cell(excelRowIndex, 40).Value = udiDi;
                                worksheet.Cell(excelRowIndex, 41).Value = status;
                                worksheet.Cell(excelRowIndex, 42).Value = secondaryUdi;
                                worksheet.Cell(excelRowIndex, 43).Value = nomenclatureCode;
                                worksheet.Cell(excelRowIndex, 44).Value = tradeName;
                                worksheet.Cell(excelRowIndex, 45).Value = catalogueNumber;
                                worksheet.Cell(excelRowIndex, 46).Value = directMarking;
                                worksheet.Cell(excelRowIndex, 47).Value = unitOfUseText;
                                worksheet.Cell(excelRowIndex, 48).Value = quantity;
                                worksheet.Cell(excelRowIndex, 49).Value = udiPi;
                                worksheet.Cell(excelRowIndex, 50).Value = additionalDescription;
                                worksheet.Cell(excelRowIndex, 51).Value = infoUrl;
                                worksheet.Cell(excelRowIndex, 52).Value = clinicalSizes;
                                worksheet.Cell(excelRowIndex, 53).Value = singleUse;
                                worksheet.Cell(excelRowIndex, 54).Value = maxNoReusesText;
                                worksheet.Cell(excelRowIndex, 55).Value = sterilisation;
                                worksheet.Cell(excelRowIndex, 56).Value = sterile;
                                worksheet.Cell(excelRowIndex, 57).Value = latex;
                                worksheet.Cell(excelRowIndex, 58).Value = handlingCondText;
                                worksheet.Cell(excelRowIndex, 59).Value = warnings;
                                worksheet.Cell(excelRowIndex, 60).Value = reprocessed;
                                worksheet.Cell(excelRowIndex, 61).Value = intendedPurpose;
                                worksheet.Cell(excelRowIndex, 62).Value = memberState;
                                worksheet.Cell(excelRowIndex, 63).Value = medProdText;

                                //
                                ////Market distribution
                                //
                                worksheet.Cell(excelRowIndex, 64).Value = versionText4;
                                worksheet.Cell(excelRowIndex, 65).Value = lastUpdateText4;
                                worksheet.Cell(excelRowIndex, 66).Value = memberStateAvailab;
                                worksheet.Cell(excelRowIndex, 67).Value = refNoText;
                                worksheet.Cell(excelRowIndex, 68).Value = revNoText;
                                worksheet.Cell(excelRowIndex, 69).Value = issueDateText;
                                worksheet.Cell(excelRowIndex, 70).Value = certificateNoText;

                                worksheet.Cell(excelRowIndex, 1).Value = udiDi;


                                Console.WriteLine($"*****************************************************************Datasaved in row {excelRowIndex}");
                                excelRowIndex++;



                                // Go back to the previous page
                                Console.WriteLine("Navigating back to the previous page...");
                                driver.Navigate().Back();

                                // Save the Excel file
                                Console.WriteLine("Saving the extracted data to an Excel file...");
                                workbook.SaveAs("Eudamed_Device_Data_2209.xlsx");

                                Console.WriteLine($"Data extraction for a product No {i + 1}! Excel file saved as 'Eudamed_Device_Data_2209.xlsx'.");


                                // Wait for the table to reload
                                Console.WriteLine("Waiting for the table to reload...");
                                Thread.Sleep(5000); // Adjust as needed


                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error occurred while processing product {i + 1}: {ex.Message}");

                                // Prompt user for action
                                Console.WriteLine("Choose an option:");
                                Console.WriteLine("[R] Retry this iteration");
                                Console.WriteLine("[S] Skip to next iteration");
                                Console.WriteLine("[E] Exit program");

                                string userChoice = Console.ReadLine()?.ToUpper();
                                switch (userChoice)
                                {
                                    case "R":
                                        retryIteration = true; // Repeat this iteration
                                        Console.WriteLine("Retrying...");
                                        break;
                                    case "S":
                                        Console.WriteLine("Skipping to the next product...");
                                        break; // Moves to next iteration
                                    case "E":
                                        Console.WriteLine("Exiting program...");
                                        driver.Quit();
                                        return; // Exit function
                                    default:
                                        Console.WriteLine("Invalid choice. Skipping iteration.");
                                        break;
                                }
                            }







                            Console.WriteLine($"Moving to page {currentPage + 1}...");
                            NavigateToNextPage((ChromeDriver)driver, currentPage);
                            // Wait until table rows are visible
                            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector("table tbody tr")));

                        } while (retryIteration);



                }
            }

            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);

                // Prompt user for action
                Console.WriteLine("Choose an option:");
                Console.WriteLine("[R] Retry this iteration");
                Console.WriteLine("[S] Skip to next iteration");
                Console.WriteLine("[E] Exit program");

                string userChoice = Console.ReadLine()?.ToUpper();
                switch (userChoice)
                {
                    case "R":
                        retryIteration = true; // Repeat this iteration
                        Console.WriteLine("Retrying...");
                        break;
                    case "S":
                        Console.WriteLine("Skipping to the next product...");
                        break; // Moves to next iteration
                    case "E":
                        Console.WriteLine("Exiting program...");
                        

                        return; // Exit function
                    default:
                        Console.WriteLine("Invalid choice. Skipping iteration.");
                        break;
                }
            }
            

            // Hold the application open until manually closed
            Console.WriteLine("Press Enter to exit the application.");
            Console.ReadLine();


        }
        // Navigate to the next page
        public static void NavigateToNextPage(ChromeDriver driver, int currentPage)
        {

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));

            // Find and click the next page button
            var nextPageButtonXPath = $"//button[@aria-label='Page number {currentPage + 1} ']";
            var nextPageButton = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(nextPageButtonXPath)));
            ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", nextPageButton);
            nextPageButton.Click();

            // Wait for the table to update
            wait.Until(d =>
            {
                var table = d.FindElement(By.TagName("p-table"));
                var rows = table.FindElements(By.CssSelector("tbody > tr"));
                return rows.Count > 0; // Ensure rows are loaded
            });

            Console.WriteLine($"Page {currentPage + 1} loaded.");
        }
        static void ClickPage(IWebDriver driver, int pageNumber)
        {
            try
            {
                IWebElement pageButton = driver.FindElement(By.XPath($"//button[contains(@aria-label, 'Page number {pageNumber}')]"));
                pageButton.Click();
                Console.WriteLine($"Navigated to Page {pageNumber}");
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine($"Page {pageNumber} button not found!");
            }
        }
    }
}

