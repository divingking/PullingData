using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace PullingData
{
    class Program
    {
        private static DataReader reader;
        private static StreamWriter file;
        private static StreamWriter error;

        /*static void Main(string[] args)
        {
            //AddBeneficiary();
            //RecoverBeneficiaryVoucher();
            //FixName();
            //FixVoucherInvoice();

            //ProcessTotalFile(@"E:\Work\WorldConcern\PullingData\PullingData\bin\Debug\Beneficiary list master - 112512.xls");

            Console.WriteLine("Full domain name: ");
            //Console.WriteLine(System.Net.Dns.GetHostName());
            Console.WriteLine(System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName);


            StringBuilder usernameWithFqdn = new StringBuilder(1024);
            int usernameWithFqdnSize = usernameWithFqdn.Capacity;
            int err = GetUserNameEx(12, usernameWithFqdn, ref usernameWithFqdnSize);

            if (usernameWithFqdn.Length == 0)
            {
                Console.WriteLine("hahahahahaha");
            }
            else
            {
                Console.WriteLine(usernameWithFqdn.ToString());
            }
        }*/

        static void Main(string[] args)
        {
            file = new StreamWriter(@"E:\Work\WorldConcern\PullingData\PullingData\bin\Debug\suppliesProduct.txt");

            string[] upcs = new string[] 
            { 
                "TBED_MIETSB11", 
                "TBED_MIETSB31", 
                "TBED_MIETHB11", 
                "TBED_MIETHB31", 
                "TBED_MIETLB11", 
                "TBED_MIESTB31", 
                "TBED_MIEOTB11", 
                "TBED_99TRRTDT", 
                "TBED_99TDRTDT", 
                "TBED_MIEOTB31", 
                "TBED_MIETBB8T", 
                "TBED_MIEBDB8T"
            };

            foreach (string upc in upcs)
            {
                for (int i = 1; i <= 26; i++)
                {
                    file.WriteLine("insert into Supplies values({0}, '{1}', 0);", i, upc);
                }
            }

            file.Flush();
            file.Close();
        }


        /// <summary>
        /// Imported Windows API to get user name
        /// </summary>
        /// <param name="nameFormat">Currently only passed in 12, which is the format code for user domain name</param>
        /// <param name="username">For store the output string</param>
        /// <param name="userNameSize">Maximum size of the returned string</param>
        /// <returns>Return code of this API</returns>
        [DllImport("Secur32.dll", CharSet = CharSet.Auto)]
        internal static extern int GetUserNameEx(int nameFormat, StringBuilder username, ref int userNameSize);

        private static void ProcessTotalFile(string inputFileName)
        {
            reader = new DataReader(inputFileName);
            file = new StreamWriter(@"E:\Work\WorldConcern\PullingData\PullingData\bin\Debug\TotalBeneficiaryVoucherTest.sql");
            try
            {
                while (reader.hasNextSheet())
                {
                    reader.nextSheet();
                    if (reader.getCurrentSheetName().Equals("Total"))
                    {
                        int id = 2;
                        while (reader.hasNextRow())
                        {
                            reader.nextRow();
                            if (isData())
                            {
                                ProcessTotalRow(id);
                                id++;
                            }
                        }

                        break;
                    }
                }
            }
            finally
            {
                reader.close();
                file.Flush();
                file.Close();
            }
        }

        private static void ProcessTotalRow(int id)
        {
            string CycleCity = reader.nextCol();
            string type = reader.nextCol();
            int typeid = 0;
            if (type.Equals("IDP"))
            {
                typeid = 1;
            }
            string name = reader.nextCol();
            string[] names = GetName(name);
            string sex = reader.nextCol();
            int age = 2012 - reader.nextInt();
            string phone = reader.nextCol();
            string hometown = reader.nextCol();
            int u5m = reader.nextInt();
            int u5f = reader.nextInt();
            int cm = reader.nextInt();
            int cf = reader.nextInt();
            int am = reader.nextInt();
            int af = reader.nextInt();
            int familySize = reader.nextInt();
            int meals = reader.nextInt();
            bool disabled = reader.nextBool();
            bool orphan = reader.nextBool();
            bool lostLiveStock = reader.nextBool();
            bool femaleHeaded = reader.nextBool();

            int v1 = reader.nextInt();
            int v2 = reader.nextInt();
            int v3 = reader.nextInt();

            file.Write("insert into Beneficiary values(");
            printName(name);
            if (typeid == 0)
            {
                file.Write(GetCity(CycleCity.Trim()) + ", ");
            }
            else
            {
                file.Write(GetCity(hometown.Trim()) + ", ");
            }
            file.Write(familySize + ", ");
            file.Write(typeid + ", ");
            file.Write(age + ", ");
            printBool(femaleHeaded);
            file.Write(u5f + ", ");
            file.Write(u5m + ", ");
            file.Write(cf + ", ");
            file.Write(cm + ", ");
            file.Write(af + ", ");
            file.Write(am + ", ");
            printBool(disabled);
            printBool(orphan);
            printBool(lostLiveStock);
            file.WriteLine("CONVERT(BIGINT, LEFT(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(30),CURRENT_TIMESTAMP,126),'-',''),':',''),'T',''),14)));");

            if (v1 != 0)
            {
                file.WriteLine("update Voucher set beneficiary = " + id + " where id = " + v1 + "; ");
            }
            if (v2 != 0)
            {
                file.WriteLine("update Voucher set beneficiary = " + id + " where id = " + v2 + "; ");
            }
            if (v3 != 0)
            {
                file.WriteLine("update Voucher set beneficiary = " + id + " where id = " + v3 + "; ");
            }
        }

        private static void FixVoucherInvoice()
        {
            StreamReader input = new StreamReader(@"E:\Work\WorldConcern\PullingData\PullingData\bin\Debug\vouchers.txt");
            file = new StreamWriter("FixVoucherInvoice.sql");
            //1	Rice	0.873	25
            //2	Beans	0.757	10
            //3	Oils	5.565	3
            //4	Salt	0.447	0.5
            //5	Sugar	1.022	1.5
            //6	WheatFlour	1.056	10

            int invoice = 0;

            while (input.Peek() > 0)
            {
                string line = input.ReadLine();
                if (line.Length > 0)
                {
                    string[] columns = line.Split(' ');
                    if (columns.Length == 1)
                    {
                        invoice = Convert.ToInt32(columns[0]);
                    }
                    else
                    {
                        int voucher = Convert.ToInt32(columns[2]);
                        file.WriteLine("update Voucher set invoice = " + invoice + " where id = " + voucher + ";");
                        file.WriteLine("insert into Consists values(" + voucher + ", 1, 0.873, 25);");
                        file.WriteLine("insert into Consists values(" + voucher + ", 2, 0.757, 10);");
                        file.WriteLine("insert into Consists values(" + voucher + ", 3, 5.565, 3);");
                        file.WriteLine("insert into Consists values(" + voucher + ", 4, 0.447, 0.5);");
                        file.WriteLine("insert into Consists values(" + voucher + ", 5, 1.022, 1.5);");
                        file.WriteLine("insert into Consists values(" + voucher + ", 6, 1.056, 10);");
                    }
                }
            }

            file.Flush();
            file.Close();
        }

        private static void FixName()
        {
            StreamReader input = new StreamReader(@"E:\Work\WorldConcern\PullingData\PullingData\bin\Debug\name-with-space.txt");
            file = new StreamWriter("FixName.sql");

            while (input.Peek() > 0)
            {
                string line = input.ReadLine();
                string[] columns = line.Split(' ');

                int id = Convert.ToInt32(columns[0]);
                string lname = columns[1].Trim();
                for (int i = 2; i < columns.Length; i++)
                {
                    lname += (" " + columns[i].Trim());
                }

                Console.WriteLine(lname);

                file.WriteLine("update Beneficiary set lname='" + lname + "' where id = " + id + ";");
            }

            input.Close();

            file.Flush();
            file.Close();
        }

        private static void AddBeneficiary()
        {
            string[] inputFiles = Directory.GetFiles(@"E:\Work\WorldConcern\Backup", "*.xls*");

            file = new StreamWriter(@"E:\Work\WorldConcern\TotalBeneficiary.sql");

            for (int i = 0; i < inputFiles.Length; i++)
            {
                Console.WriteLine("Processing " + inputFiles[i]);
                processBeneficiaryFile(inputFiles[i], GetCityName(inputFiles[i]));
                Console.WriteLine("Done processing " + inputFiles[i]);
            }

            file.Flush();
            file.Close();
        }

        private static void RecoverBeneficiaryVoucher()
        {
            string[] inputFiles = Directory.GetFiles(@"E:\Downloads\Update on scanning  & feedbacks", "*.xls");

            file = new StreamWriter(@"E:\Work\WorldConcern\BeneficiaryVoucherRecovery.sql");
            error = new StreamWriter(@"E:\Work\WorldConcern\MissingBeneficiary.txt");

            for (int i = 0; i < inputFiles.Length; i++)
            {
                Console.WriteLine("Processing " + inputFiles[i]);
                error.WriteLine("Processing " + inputFiles[i]);
                processBeneficiaryVoucherFile(inputFiles[i]);
                Console.WriteLine("Done processing " + inputFiles[i]);
                error.WriteLine("Done processing " + inputFiles[i]);
            }

            file.Flush();
            file.Close();
            error.Flush();
            error.Close();
        }

        private static void processBeneficiaryVoucherFile(string inputFileName)
        {
            reader = new DataReader(inputFileName);
            try
            {
                while (reader.hasNextSheet())
                {
                    reader.nextSheet();
                    String SheetName = reader.getCurrentSheetName();
                    Console.WriteLine(SheetName);
                    error.WriteLine(SheetName);

                    while (reader.hasNextRow())
                    {
                        reader.nextRow();
                        if (isData())
                        {
                            processBeneficiaryVoucherRow();
                        }
                    }
                }
            }
            finally
            {
                reader.close();
            }
        }

        private static void processBeneficiaryVoucherRow()
        {
            string name = reader.nextCol();
            string[] names = GetName(name);
            if (names != null)
            {
                reader.nextCol();
                int age = reader.nextInt();
                for (int i = 0; i < 8; i++)
                {
                    reader.nextCol();
                }
                int familysize = reader.nextInt();
                reader.nextCol();
                Console.WriteLine(names[0] + " " + names[1] + " " + names[2]);
                string url = "http://worldconcerndataservice.cloudapp.net/worldconcerndataservice.svc/beneficiary_id?fname=" + names[0] + "&mname=" + names[1] + "&lname=" + names[2] + "&birth=" + (2012 - age) + "&family=" + familysize;
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream resStream = response.GetResponseStream();
                StreamReader result = new StreamReader(resStream);
                int id = Convert.ToInt32(result.ReadToEnd());
                
                int voucher = reader.nextInt();

                if (id == 0)
                {
                    Console.WriteLine("===========================");
                    Console.WriteLine("ERROR:");
                    Console.WriteLine("Name: " + names[0] + " " + names[1] + " " + names[2]);
                    Console.WriteLine("Year of birth: " + (2012 - age));
                    Console.WriteLine("Family size: " + familysize);
                    Console.WriteLine("Voucher: " + voucher);
                    Console.WriteLine("===========================");
                    error.WriteLine("===========================");
                    error.WriteLine("ERROR:");
                    error.WriteLine("Name: " + names[0] + " " + names[1] + " " + names[2]);
                    error.WriteLine("Year of birth: " + (2012 - age));
                    error.WriteLine("Family size: " + familysize);
                    error.WriteLine("Voucher: " + voucher);
                    error.WriteLine("===========================");
                }
                else
                {
                    file.WriteLine("update Voucher set beneficiary = " +  id + " where id = " + voucher + "; ");
                }
            }
        }

        private static string GetCityName(string fileName)
        {
            string[] parsed = fileName.Split(' ');
            string[] first = parsed[0].Split('\\');
            return first[first.Length - 1];
        }

        private static void processBeneficiaryFile(string inputFileName, string cityName)
        {
            reader = new DataReader(inputFileName);

            while (reader.hasNextSheet())
            {
                reader.nextSheet();
                String name = reader.getCurrentSheetName();
                Console.WriteLine(name);
                int city = -1;
                if (name.ToLower().Contains("host"))
                {
                    city = GetCity(cityName);
                }
                while (reader.hasNextRow())
                {
                    reader.nextRow();
                    if (isData())
                    {
                        processBeneficiaryRow(city);
                    }
                }
                Console.WriteLine("Next Sheet! ");
            }

            reader.close();
        }

        private static bool isData()
        {
            string first = reader.nextCol();
            int id;
            return int.TryParse(first, out id);
        }

        private static void processBeneficiaryRow(int city)
        {
            string name = reader.nextCol();
            if (name != null)
            {
                string sex = reader.nextCol();
                int age = reader.nextInt();
                reader.nextCol();
                string home = "";
                if (city == -1)
                {
                    home = reader.nextCol();
                }
                else
                {
                    reader.nextCol();
                }
                int u5m = reader.nextInt();
                int u5f = reader.nextInt();
                int cm = reader.nextInt();
                int cf = reader.nextInt();
                int am = reader.nextInt();
                int af = reader.nextInt();
                int familySize = reader.nextInt();
                int meals = reader.nextInt();
                bool disabled = reader.nextBool();
                bool orphan = reader.nextBool();
                bool lostLiveStock = reader.nextBool();
                bool femaleHeaded = reader.nextBool();
                file.Write("insert into Beneficiary values(");
                printName(name);
                if (city == -1)
                {
                    int homecode = GetCity(home.Trim());
                    if (homecode == -1)
                    {
                        Console.WriteLine(home);
                    }
                    file.Write(GetCity(home.Trim()) + ", ");
                }
                else
                {
                    file.Write(city + ", ");
                }
                file.Write(familySize + ", ");
                if (city == -1)
                {
                    file.Write(1 + ", ");
                }
                else
                {
                    file.Write(0 + ", ");
                }
                file.Write((2012 - age) + ", ");
                printBool(femaleHeaded);
                file.Write(u5f + ", ");
                file.Write(u5m + ", ");
                file.Write(cf + ", ");
                file.Write(cm + ", ");
                file.Write(af + ", ");
                file.Write(am + ", ");
                printBool(disabled);
                printBool(orphan);
                printBool(lostLiveStock);
                file.WriteLine("CONVERT(BIGINT, LEFT(REPLACE(REPLACE(REPLACE(CONVERT(VARCHAR(30),CURRENT_TIMESTAMP,126),'-',''),':',''),'T',''),14)));");
            }
        }

        private static int GetCity(string city)
        {
            switch (city.ToLower())
            {
                case "dhobley":
                    return 1;
                case "degalema":
                    return 2;
                case "shaabah":
                    return 3;
                case "afmadow":
                    return 4;
                case "gadudey":
                    return 5;
                case "kamdada":
                    return 6;
                case "jilib":
                    return 7;
                case "kismayo":
                    return 8;
                case "jamame":
                    return 9;
                case "kamsuma":
                    return 10;
                case "hosingo":
                    return 11;
                case "hosingow":
                    return 11;
                case "kanjaron":
                    return 12;
                case "b/xani":
                    return 13;
                case "b/xaji":
                    return 14;
                case "baqdad":
                    return 15;
                case "buale":
                    return 16;
                case "godeya":
                    return 17;
                case "yeya":
                    return 18;
                case "jamar":
                    return 19;
                case "dudunta":
                    return 20;
                case "jira":
                    return 21;
                case "marerey":
                    return 22;
                case "tabta":
                    return 23;
                case "qoqani":
                    return 24;
                case "badhade":
                    return 25;
                case "magar":
                    return 26;
                case "yontoy":
                    return 27;
                case "hagar":
                    return 28;
                case "galaf":
                    return 29;
                case "hayo":
                    return 30;
                case "bualle":
                    return 31;
                case "badade":
                    return 32;
                case "janaabdalla":
                    return 33;
                case "jilalow":
                    return 34;
                case "nasiriya":
                    return 35;
                case "hargeisayarey":
                    return 36;
                case "bula gadud":
                    return 37;
                case "bula banan":
                    return 38;
                case "bulahaji":
                    return 39;
                case "xayo":
                    return 40;
                case "godaya":
                    return 41;
                case "bananey":
                    return 42;
                case "xagar":
                    return 43; 
                case "soya":
                    return 44;
                case "diff":
                    return 45;
                case "kokani": 
                    return 46;
                case "alibuley": 
                    return 47;
                case "kismayu": 
                    return 48;
                case "kaimah": 
                    return 49;
                case "sakow": 
                    return 50;
                case "salagle": 
                    return 51;
                case "baladhawo": 
                    return 52;
                case "dinsor": 
                    return 53;
                case "dinsoor": 
                    return 54;
                case "mido":
                    return 55;
                case "jana'able": 
                    return 56;
                case "middo": 
                    return 57;
                case "jana'abdle": 
                    return 58;
                default:
                    return -1;
            }
        }

        private static string[] GetName(string name)
        {
            string[] names = name.Replace("'", ",").Trim().Split(' ');
            if (names.Length == 0)
            {
                return null;
            }
            else
            {
                string[] standardName = new string[3];
                if (names.Length == 1)
                {
                    standardName[0] = names[0].Trim();
                    standardName[1] = "";
                    standardName[2] = "";
                }
                else if (names.Length == 2)
                {
                    standardName[0] = names[0].Trim();
                    standardName[1] = "";
                    standardName[2] = names[1].Trim();
                }
                else if (names.Length == 3)
                {
                    standardName[0] = names[0].Trim();
                    standardName[1] = names[1].Trim();
                    standardName[2] = names[2].Trim();
                }
                else
                {
                    standardName[0] = names[0].Trim();
                    standardName[1] = names[1].Trim();
                    string lastname = names[2].Trim();
                    for (int i = 3; i < names.Length; i++)
                    {
                        lastname += (" " + names[i].Trim());
                    }
                    standardName[2] = lastname.Trim();
                }
                return standardName;
            }
        }

        private static void printName(string name)
        {
            string[] names = name.Replace("'", ",").Trim().Split(' ');
            if (names.Length == 3)
            {
                file.Write("'" + names[0].Trim() + "', ");
                file.Write("'" + names[1].Trim() + "', ");
                file.Write("'" + names[2].Trim() + "', ");
            }
            else if (names.Length == 2)
            {
                file.Write("'" + names[0].Trim() + "', ");
                file.Write("'', ");
                file.Write("'" + names[1].Trim() + "', ");
            }
            else if (names.Length == 1)
            {
                file.Write("'" + names[0].Trim() + "', ");
                file.Write("'', ");
                file.Write("'', ");
            }
            else
            {
                file.Write("'" + names[0].Trim() + "', ");
                file.Write("'" + names[1].Trim() + "', ");
                file.Write("'" + names[2].Trim());
                for (int i = 3; i < names.Length; i++)
                {
                    file.Write(" " + names[i].Trim());
                }
                file.Write("', ");
            }
        }

        private static void printBool(bool b)
        {
            if (b)
            {
                file.Write(1 + ", ");
            }
            else
            {
                file.Write(0 + ", ");
            }
        }
    }
}
