using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml;
using System.Xml.Serialization;

namespace MonthlyEStatements
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Copy the Statements extract to " + Configuration.GetSymitarStatementDataFilePath() + " then press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Copy the MoneyPerks extract to " + Configuration.GetMoneyPerksFilePath() + " then press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Ensure the following path exists: " + Configuration.GetStatementsOutputPath() + " then press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Ensure the following path exists: " + Configuration.GetErrorLogOutputPath() + " then press any key to continue.");
            Console.ReadLine();
            Stopwatch stopwatch = Stopwatch.StartNew();
            LogWriter = new StreamWriter(Configuration.GetErrorLogOutputPath() + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".log");
            CleanStatementsOutputPath();
            BuildMoneyPerksStatements();
            BuildMemberStatements();
            LogWriter.Close();
            stopwatch.Stop();
            Console.WriteLine(StatementBuilder.GetNumberOfStatementsBuilt() + " statements produced in " + stopwatch.Elapsed.TotalMinutes.ToString("N") + " minutes.");
            Console.ReadLine();
        }

        public static void BuildMoneyPerksStatements()
        {
            Console.WriteLine("Processing MoneyPerks file " + Configuration.GetMoneyPerksFilePath() + "...");
            StreamReader reader = new StreamReader(Configuration.GetMoneyPerksFilePath());
            MoneyPerksStatements = new Dictionary<string, MoneyPerksStatement>();

            while (!reader.EndOfStream)
            {
                List<string> fields = CSVParser(reader.ReadLine());

                if (fields.Count == MONEYPERKS_TRANSACTION_RECORD_FIELD_COUNT)
                {
                    MoneyPerksStatement moneyPerksStatement = null;
                    MoneyPerksTransaction transaction = new MoneyPerksTransaction();
                    string accountNumber = fields[0];
                    DateTime date = ParseDate(fields[1]);
                    string description = fields[2];
                    int amount = ParseMoneyPerksAmount(fields[3]);
                    int balance = ParseMoneyPerksAmount(fields[4]);

                    if (MoneyPerksStatements.ContainsKey(accountNumber))
                    {
                        moneyPerksStatement = MoneyPerksStatements[accountNumber];
                    }
                    else
                    {
                        moneyPerksStatement = new MoneyPerksStatement(accountNumber);
                        MoneyPerksStatements.Add(accountNumber, moneyPerksStatement);
                    }

                    if (description == "Beginning Balance")
                    {
                        moneyPerksStatement.BeginningBalance = balance;
                    }
                    else if (description == "Ending Balance")
                    {
                        moneyPerksStatement.EndingBalance = balance;
                    }
                    else
                    {
                        transaction.Date = date;
                        transaction.Description = description;
                        transaction.Amount = amount;
                        transaction.Balance = balance;
                        moneyPerksStatement.Transactions.Add(transaction);
                        MoneyPerksStatements[accountNumber] = moneyPerksStatement;
                    }
                }
            }

            reader.Close();
            reader.Dispose();
            Console.WriteLine("Done processing MoneyPerks file");

        }

        static void CleanStatementsOutputPath()
        {
            Console.WriteLine("I am going to delete all of the files in " + Configuration.GetStatementsOutputPath() + ".  Press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Deleting files from " + Configuration.GetStatementsOutputPath() + "...");
            System.IO.DirectoryInfo dirInfo = new DirectoryInfo(Configuration.GetStatementsOutputPath());

            foreach (FileInfo fileInfo in dirInfo.GetFiles())
            {
                fileInfo.Delete();
            }
        }

        static void BuildMemberStatements()
        {
            Console.WriteLine("Processing Statement file " + Configuration.GetSymitarStatementDataFilePath() + "...");
            
            //string accountNumber = MemberStatement.envelope[0].statement[0].account[0].accountNumber;


            //StatementBuilder statement = new StatementBuilder();
            DeserializeObject("C:\\Samples\\XML\\TestingStatement.xml");


            
            Console.WriteLine("Done processing Statement file");
            Console.ReadLine();
        }

        public static void DeserializeObject(string filename)
        {
            MemberStatement = new StatementProduction();

            XmlSerializer serializer = new XmlSerializer(typeof(StatementProduction));

            FileStream fs = new FileStream(filename, FileMode.Open);
            XmlReader reader = new XmlTextReader(fs);

            //StatementProduction s = new StatementProduction();

            //s = (StatementProduction)serializer.Deserialize(reader);
            MemberStatement = (StatementProduction)serializer.Deserialize(reader);

 
            StatementBuilder.Build(MemberStatement, null);
            reader.Close();
        }


        static DateTime ParseDate(string date)
        {
            DateTime parsedDate = new DateTime();

            date = date.Replace("-", string.Empty); // Removes dash from MoneyPerks file dates

            if (date.Length >= "MMDDYYYY".Length)
            {
                try
                {
                    int year = int.Parse(date.Substring("MMDD".Length, "YYYY".Length));
                    int month = int.Parse(date.Substring(0, "MM".Length));
                    int day = int.Parse(date.Substring("MM".Length, "DD".Length));
                    parsedDate = new DateTime(year, month, day);
                }
                catch (Exception exception)
                {
                    Log(exception.Message);
                }
            }

            return parsedDate;
        }

        private static int ParseMoneyPerksAmount(string amount)
        {
            int parsedAmount = 0;

            if (amount != string.Empty)
            {
                try
                {
                    parsedAmount = int.Parse(amount);
                }
                catch (Exception exception)
                {
                    Log(exception.Message + " " + amount);
                }
            }

            return parsedAmount;
        }

        private static List<string> CSVParser(string strInputString)
        {
            int intCounter = 0, intLenght;
            StringBuilder strElem = new StringBuilder();
            List<string> alParsedCsv = new List<string>();
            intLenght = strInputString.Length;
            strElem = strElem.Append("");
            int intCurrState = 0;
            int[][] aActionDecider = new int[9][];
            //Build the state array
            aActionDecider[0] = new int[4] { 2, 0, 1, 5 };
            aActionDecider[1] = new int[4] { 6, 0, 1, 5 };
            aActionDecider[2] = new int[4] { 4, 3, 3, 6 };
            aActionDecider[3] = new int[4] { 4, 3, 3, 6 };
            aActionDecider[4] = new int[4] { 2, 8, 6, 7 };
            aActionDecider[5] = new int[4] { 5, 5, 5, 5 };
            aActionDecider[6] = new int[4] { 6, 6, 6, 6 };
            aActionDecider[7] = new int[4] { 5, 5, 5, 5 };
            aActionDecider[8] = new int[4] { 0, 0, 0, 0 };
            for (intCounter = 0; intCounter < intLenght; intCounter++)
            {
                intCurrState = aActionDecider[intCurrState]
                                          [CSVParser_GetInputID(strInputString[intCounter])];
                //take the necessary action depending upon the state 
                CSVParser_PerformAction(ref intCurrState, strInputString[intCounter],
                             ref strElem, ref alParsedCsv);
            }
            //End of line reached, hence input ID is 3
            intCurrState = aActionDecider[intCurrState][3];
            CSVParser_PerformAction(ref intCurrState, '\0', ref strElem, ref alParsedCsv);
            return alParsedCsv;
        }

        private static int CSVParser_GetInputID(char chrInput)
        {
            if (chrInput == '"')
            {
                return 0;
            }
            else if (chrInput == ',')
            {
                return 1;
            }
            else
            {
                return 2;
            }
        }
        private static void CSVParser_PerformAction(ref int intCurrState, char chrInputChar,
                            ref StringBuilder strElem, ref List<string> alParsedCsv)
        {
            string strTemp = null;
            switch (intCurrState)
            {
                case 0:
                    //Separate out value to array list
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    strElem = new StringBuilder();
                    break;
                case 1:
                case 3:
                case 4:
                    //accumulate the character
                    strElem.Append(chrInputChar);
                    break;
                case 5:
                    //End of line reached. Separate out value to array list
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    break;
                case 6:
                    //Erroneous input. Reject line.
                    alParsedCsv.Clear();
                    break;
                case 7:
                    //wipe ending " and Separate out value to array list
                    strElem.Remove(strElem.Length - 1, 1);
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    strElem = new StringBuilder();
                    intCurrState = 5;
                    break;
                case 8:
                    //wipe ending " and Separate out value to array list
                    strElem.Remove(strElem.Length - 1, 1);
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    strElem = new StringBuilder();
                    //goto state 0
                    intCurrState = 0;
                    break;
            }
        }

        private static void Log(string message)
        {
            try
            {
                StackTrace stackTrace = new StackTrace();
                StackFrame stackFrame = stackTrace.GetFrame(1);
                MethodBase methodBase = stackFrame.GetMethod();
                string methodName = methodBase.Name;

                if (MemberStatement != null)
                {
                   // LogWriter.WriteLine(methodBase.DeclaringType.Name + "." + methodName + ": " + message + " (Account Number: " + MemberStatement.envelope[0].statement.account.accountNumber + ")");
                }
                else
                {
                    LogWriter.WriteLine(methodBase.DeclaringType.Name + "." + methodName + ": " + message);
                }
            }

            catch (Exception)
            {
            }
        }

        static StreamWriter LogWriter
        {
            get;
            set;
        }

        static string AccountNumber
        {
            get;
            set;
        }

        static StatementProduction MemberStatement
        {
            get;
            set;
        }

        static Dictionary<string, MoneyPerksStatement> MoneyPerksStatements
        {
            get;
            set;
        }

       

        public const int MONEYPERKS_TRANSACTION_RECORD_FIELD_COUNT = 5;
      
    }
}
