using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml.Serialization;
using System.Xml;
using MonthlyEStatements.Common;


namespace MonthlyEStatements
{
    public class StatementBuilder
    {
        
        public static void Build(StatementProduction statement, MoneyPerksStatement moneyPerksStatement)
        {

            AdvertisementTop = new Advertisement[MAX_RELATIONSHIP_BASED_LEVELS];

            for (int i = 0; i < AdvertisementTop.Count(); i++)
            {
                AdvertisementTop[i] = new Advertisement();
            }

            using (FileStream outputStream = File.Create("C:\\" + TEMP_FILE_NAME))
            {
                CreateFirstPage(statement, outputStream);
                Share share;
                Loan loan;
                Account[] accounts;
                SubAccount[] subAccounts;
                Statement[] currentStatement;

                //Step through envelopes
                for (int i = 0; i < statement.envelope.Count();i++ )
                {
                    currentStatement = statement.envelope[i].statement;
                    //Step through statements
                    for(int j = 0; j < currentStatement.Count(); j++)
                    {
                        accounts = statement.envelope[i].statement[j].account;
                        //Step through accounts
                        for(int m = 0; m < accounts.Count(); m++)
                        {
                            subAccounts = accounts[m].subAccount;
                            //Step through subaccounts
                            for(int p = 0; p < subAccounts.Count(); p++)
                            {
                                string item = subAccounts[p].Item.ToString();
                                
                                switch(item)
                                {
                                    case "Share":
                                        share = (Share)subAccounts[p].Item;
                                        AddShare(share);
                                        break;
                                    case "Loan":
                                        loan = (Loan)subAccounts[p].Item;
                                        AddLoan(loan);
                                        break;
                                    default:
                                        Console.WriteLine("Invalid");
                                        break;
                                }
                            }
                        }
                    }
                }


                /*
                    for (int i = 0; i < statement.envelope.Count(); i++)
                    {
                        currentStatement = statement.envelope[i].statement;
                        accounts = statement.envelope[i].statement[0].account;

                        for (int j = 0; j < accounts.Count(); j++)
                        {
                            subAccounts = accounts[j].subAccount;

                            for (int m = 0; m < subAccounts.Count(); m++)
                            {
                                string item = statement.envelope[i].statement[0].account[j].subAccount[m].Item.ToString();

                                if ((item) == "Share")
                                {
                                    share = (Share)statement.envelope[i].statement[0].account[j].subAccount[m].Item;
                                    AddShare(share);

                                }
                                else if ((item) == "Loan")
                                {
                                    loan = (Loan)statement.envelope[i].statement[0].account[j].subAccount[m].Item;
                                    AddLoan(loan);
                                }
                                else
                                {
                                    Console.WriteLine("Invalid");
                                }
                            }

                        }
                    }
                */
                //AddYtdSummaries(share, loan);
                AddMoneyPerksSummary(moneyPerksStatement);
                AddBottomAdvertising(statement);
                Doc.Close();
            }

            AddPageNumbersAndDisclosures(statement); // Re-opens document to overlay page numbers

            NumberOfStatementsBuilt++;
        }


        static void AddShare(Share share)
        {
            string value = share.category.Value;
            
            switch(value)
            {
                case "Share":
                    AddSavingsAccounts(share);
                    break;
                case "Draft":
                    AddCheckingAccounts(share);
                    break;
                case "Club":
                    AddClubAccounts(share);
                    break;
                case "Certificate":
                    AddCertificateAccounts(share);
                    break;
                default:
                    break;
            }

        }

        static void AddCheckingAccounts(Share draft)
        {
            int i = 1;
            decimal annualRate = 0;
            DateTime apyePeriodStartDate = DateTime.Now;
            DateTime apyePeriodEndDate = DateTime.Now;
            decimal overdraftFeeYTD = 0;
            decimal totalReturnedItemFeeYTD = 0;

            AddSectionHeading("CHECKING ACCOUNTS");
            AddAccountSubHeading(draft.description, i > 0);
            AddAccountTransactions(draft);

            if(draft.transaction != null)
            {
                Transaction[] draftTransactions = draft.transaction;

                annualRate = Convert.ToDecimal(draftTransactions[0].apyeRate);
                apyePeriodStartDate = Convert.ToDateTime(draftTransactions[0].apyePeriodStartDate);
                apyePeriodEndDate = Convert.ToDateTime(draftTransactions[0].apyePeriodEndDate);
            }

            if(draft.overdraftFeeYTD != null)
            {
                overdraftFeeYTD = Convert.ToDecimal(draft.overdraftFeeYTD);
            }

            if(draft.returnedItemFeeYTD != null)
            {
                totalReturnedItemFeeYTD = Convert.ToDecimal(draft.returnedItemFeeYTD);
            }
            

                // Adds APR
                if (annualRate > 0)
                {
                    PdfPTable table = new PdfPTable(1);
                    table.TotalWidth = 525f;
                    table.LockedWidth = true;
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Annual Percentage Yield Earned " + annualRate.ToString("N3") + "% from " + apyePeriodStartDate.ToString("MM/dd/yyyy") + " through " + apyePeriodEndDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.IndentationLeft = 70;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                    Doc.Add(table);
                }
            /*
                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (statement.Checks.Count > 0)
                {
                    AddChecks(statement);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }

            */
                if ((overdraftFeeYTD + totalReturnedItemFeeYTD) > 0)
                {
                    AddTotalFees(draft);
                }
        }

        static void AddSavingsAccounts(Share savings)
        {
            decimal annualRate = 0;
            decimal overdraftFeeYTD = 0;
            decimal totalReturnedItemFeeYTD = 0;
            DateTime apyeStartDate = DateTime.Now;
            DateTime apyeEndDate = DateTime.Now;
            int i = 1;


            AddSectionHeading("SAVINGS ACCOUNTS");
            AddAccountSubHeading(savings.description, i > 0);
            AddAccountTransactions(savings);
            
            if(savings.overdraftFeeYTD != null)
            {
                overdraftFeeYTD = Convert.ToDecimal(savings.overdraftFeeYTD);
            }
            if(savings.returnedItemFeeYTD != null)
            {
                totalReturnedItemFeeYTD = Convert.ToDecimal(savings.returnedItemFeeYTD);
            }

            if(savings.transaction != null)
            {
                Transaction[] savingsTransactions = savings.transaction;

                annualRate = Convert.ToDecimal(savings.transaction[0].apyeRate);
                apyeStartDate = Convert.ToDateTime(savings.transaction[0].apyePeriodStartDate);
                apyeEndDate = Convert.ToDateTime(savings.transaction[0].apyePeriodEndDate);
                
            }

                // Adds APR
                if (annualRate > 0)
                {
                    PdfPTable table = new PdfPTable(1);
                    table.TotalWidth = 525f;
                    table.LockedWidth = true;
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Annual Percentage Yield Earned " + annualRate.ToString("N3") + "% from " + apyeStartDate.ToString("MM/dd/yyyy") + " through " + apyeEndDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.IndentationLeft = 70;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                    Doc.Add(table);
                }
            /*
                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }
            */
                if ((overdraftFeeYTD + totalReturnedItemFeeYTD) > 0)
                {
                    AddTotalFees(savings);
                }
            
        }

        static void AddClubAccounts(Share club)
        {
            DateTime apyeStartDate = DateTime.Now;
            DateTime apyeEndDate = DateTime.Now;
            decimal annualRate = 0;
            decimal overdraftFeeYTD = 0;
            decimal totalReturnedItemFeeYTD = 0;
            int i = 1;
            
            AddSectionHeading("CLUB ACCOUNTS");
            AddAccountSubHeading(club.description, i > 0);
            AddAccountTransactions(club);

            if (club.overdraftFeeYTD != null)
            {
                overdraftFeeYTD = Convert.ToDecimal(club.overdraftFeeYTD);
            }
            if (club.returnedItemFeeYTD != null)
            {
                totalReturnedItemFeeYTD = Convert.ToDecimal(club.returnedItemFeeYTD);
            }

            if (club.transaction != null)
            {
                Transaction[] savingsTransactions = club.transaction;

                annualRate = Convert.ToDecimal(club.transaction[0].apyeRate);
                apyeStartDate = Convert.ToDateTime(club.transaction[0].apyePeriodStartDate);
                apyeEndDate = Convert.ToDateTime(club.transaction[0].apyePeriodEndDate);

            }

                // Adds APR
                if (annualRate > 0)
                {
                    PdfPTable table = new PdfPTable(1);
                    table.TotalWidth = 525f;
                    table.LockedWidth = true;
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Annual Percentage Yield Earned " + annualRate.ToString("N3") + "% from " + apyeStartDate.ToString("MM/dd/yyyy") + " through " + apyeEndDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.IndentationLeft = 70;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                    Doc.Add(table);
                }
/*
                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }

*/
                if ((overdraftFeeYTD + totalReturnedItemFeeYTD) > 0)
                {
                    AddTotalFees(club);
                }
            
        }

        static void AddCertificateAccounts(Share certificate)
        {
            DateTime maturityDate = Convert.ToDateTime(certificate.maturityDate);
            DateTime apyeStartDate = DateTime.Now;
            DateTime apyeEndDate = DateTime.Now;
            int i = 1;
            decimal overdraftFeeYTD = 0;
            decimal totalReturnedItemFeeYTD = 0;
            decimal annualRate = 0;


            AddSectionHeading("CERTIFICATE ACCOUNTS");
            string descriptionAndMaturityDate = certificate.description + "   Maturity Date - " + maturityDate.ToString("MMM dd, yyyy");
            AddAccountSubHeading(descriptionAndMaturityDate, i > 0);


                if (certificate.overdraftFeeYTD != null)
                {
                    overdraftFeeYTD = Convert.ToDecimal(certificate.overdraftFeeYTD);
                }
                if (certificate.returnedItemFeeYTD != null)
                {
                    totalReturnedItemFeeYTD = Convert.ToDecimal(certificate.returnedItemFeeYTD);
                }

                if (certificate.transaction != null)
                {
                    Transaction[] savingsTransactions = certificate.transaction;

                    annualRate = Convert.ToDecimal(certificate.transaction[0].apyeRate);
                    apyeStartDate = Convert.ToDateTime(certificate.transaction[0].apyePeriodStartDate);
                    apyeEndDate = Convert.ToDateTime(certificate.transaction[0].apyePeriodEndDate);

                }


                AddAccountTransactions(certificate);
            /*
                if (account.CheckHolds.Count() > 0)
                {
                    AddCheckHolds(account);
                }

                if (account.AtmWithdrawals.Count > 0)
                {
                    AddAtmWithdrawals(account);
                }

                if (account.AtmDeposits.Count > 0)
                {
                    AddAtmDeposits(account);
                }

            */
                if ((overdraftFeeYTD + totalReturnedItemFeeYTD) > 0)
                {
                    AddTotalFees(certificate);
                }
            
        }


        static void AddLoan(Loan loan)
        {
            decimal totalCashAdvances = 0;
            decimal totalPayments = 0; 
            DateTime endingStatementDate = Convert.ToDateTime(loan.endingStatementDate);

            AddSectionHeading("LOAN ACCOUNTS");
            int i = 1;

           AddAccountSubHeading(loan.description, i > 0);
           AddLoanPaymentInformation(loan);
           AddLoanTransactions(loan);

            if(loan.activitySummary != null)
            {
                if (loan.activitySummary.totalCashAdvances != null)
                {
                    totalCashAdvances = Convert.ToDecimal(loan.activitySummary.totalCashAdvances);
                }

                if (loan.activitySummary.totalPayments != null)
                {
                    totalPayments = Convert.ToDecimal(loan.activitySummary.totalPayments);
                }
            }
    
            if (loan.closeDate == null)
            {
                AddLoanTransactionsFooter("Closing Date of Billing Cycle " + endingStatementDate.ToString("MM/dd/yyyy") + "\n" +
                    "** INTEREST CHARGE CALCULATION: The balance used to compute interest charges is the unpaid balance each day after payments and credits to that balance have been subtracted and any additions to the balance have been made.");
                AddFeeSummary(loan);
                AddInterestChargedSummary(loan);
            }
                
            AddYearToDateTotals(loan);

        }

        static void CreateFirstPage(StatementProduction statement, FileStream outputStream)
        {
            
            DateTime statementStart = Convert.ToDateTime(statement.envelope[0].statement[0].beginningStatementDate);
            DateTime statementEnd = Convert.ToDateTime(statement.envelope[0].statement[0].endingStatementDate);
            string accountNumber = statement.envelope[0].statement[0].account[0].accountNumber;
     


                //Adds first page template to statement
                using (FileStream templateInputStream = File.Open(Configuration.GetStatementTemplateFirstPageFilePath(), FileMode.Open))
                {
                    PdfReader reader = new PdfReader(templateInputStream);
                    Doc = new Document(reader.GetPageSize(1));
                    Writer = PdfWriter.GetInstance(Doc, outputStream);
                    StatementPageEvent pageEvent = new StatementPageEvent();
                    Writer.PageEvent = pageEvent;
                    Writer.SetFullCompression();
                    Doc.Open();
                    PdfContentByte contentByte = Writer.DirectContent;
                    PdfImportedPage page = Writer.GetImportedPage(reader, 1);
                    Doc.NewPage();
                    contentByte.AddTemplate(page, 0, 0);
                }

            AddStatementHeading("Statement  of  Accounts", 409, 0);
            AddStatementHeading(statementStart.ToString("MMM  dd,  yyyy") + "  thru  " + statementEnd.ToString("MMM  dd, yyyy"), 385, 6f);

            

            //Should we iterate and add a statement heading account number for every one or just the first account #?
            //addBasicAccountDetails(statement.envelope[j].statement.account);
         

            if (accountNumber.Length > 4)
            {
                AddStatementHeading("Account  Number:        ******" + accountNumber.Substring("******".Length), 385, 6f);
            }

            AddInvisibleAccountNumber(statement);

            PdfPTable addressAndBalancesTable = new PdfPTable(2);
            float[] addressAndBalancesTableWidths = new float[] { 50f, 50f };
            addressAndBalancesTable.SetWidthPercentage(addressAndBalancesTableWidths, Doc.PageSize);
            addressAndBalancesTable.TotalWidth = 612f;
            addressAndBalancesTable.LockedWidth = true;
            AddAddress(statement, ref addressAndBalancesTable);
            AddHeaderBalances(statement, ref addressAndBalancesTable);
            Doc.Add(addressAndBalancesTable);
            AddTopAdvertising(statement);


            //Count the number of sub accounts in each envelope
            /*
            for (int i = 0; i < envelopeCount; i++)
            {
                if (statement.envelope[0].statement.account.accountNumber > 4)
                {
                    AddStatementHeading("Account  Number:     ******" + statement.envelope[0].statement.account.accountNumber, 385, 6f);
                }
                AddInvisibleAccountNumber(statement);

                subaccounts = statement.envelope[i].statement.account.subAccount;

                foreach (statementProductionEnvelopeStatementAccountSubAccount sub in subaccounts)
                {
                    subAccountCount++;
                }

                SortSubAccounts(envelopeCount, loanCounts, accountCount, subAccountCount, subaccounts, ref addressAndBalancesTable);
                i++;

                if (i < envelopeCount)
                {
                    subAccountCount = 0;
                    subaccounts = statement.envelope[i].statement.account.subAccount;

                    foreach (statementProductionEnvelopeStatementAccountSubAccount sb in subaccounts)
                    {
                        subAccountCount++;
                    }
                }

                SortSubAccounts(envelopeCount, loanCounts, accountCount, subAccountCount, subaccounts, ref addressAndBalancesTable);
            }
            */

        }

        static void AddAddress(StatementProduction statement, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(statement.envelope[0].person[0].firstName + " " + statement.envelope[0].person[0].lastName, GetNormalFont(9));
            chunk.Append("\n" + statement.envelope[0].address.street.ToUpper());
            if (statement.envelope[0].address.city != null) chunk.Append("\n" + statement.envelope[0].address.city.ToUpper() + ", " + statement.envelope[0].address.state.ToUpper());
            chunk.Append("\n" + statement.envelope[0].address.postalCode);
            
            
                //chunk.Append("\n" + additionalName.Name + ", " + additionalName.TypeString);
            

            
            chunk.setLineHeight(9f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 66;
            cell.AddElement(p);
            cell.PaddingTop = 60f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddAccountTransactions(Share share)
        {
            Transaction[] shareTransactions;

            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 280, 62, 65, 67 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            AddAccountTransactionTitle("Date", Element.ALIGN_LEFT, ref table);
            AddAccountTransactionTitle("Transaction Description", Element.ALIGN_LEFT, ref table);
            AddAccountTransactionTitle("Additions", Element.ALIGN_RIGHT, ref table);
            AddAccountTransactionTitle("Subtractions", Element.ALIGN_RIGHT, ref table);
            AddAccountTransactionTitle("Balance", Element.ALIGN_RIGHT, ref table);
            AddBalanceForward(share, ref table);

            if(share.transaction != null)
            {
                shareTransactions = share.transaction;
                for (int i = 0; i < shareTransactions.Length;i++ )
                {
                    if (shareTransactions[i].category.Value == "Comment")
                    {
                        AddCommentOnlyTransaction(shareTransactions[i], ref table);
                    }
                    else if (shareTransactions[i].category.Value == "Deposit")
                    {
                        AddDeposits(shareTransactions);
                    }
                    else if (shareTransactions[i].category.Value == "Withdrawal")
                    {
                        AddWithdrawals(shareTransactions);
                    }
                    else
                    {
                        AddAccountTransaction(shareTransactions[i], ref table);
                    }
                }
            }

            if(share.closeDate != null)
            {
                AddShareClosed(share, ref table);
            }
            else
            {
                AddEndingBalance(share, null, ref table);
            }

            Doc.Add(table);
        }

        static void AddAccountTransaction(Transaction transaction, ref PdfPTable table)
        {
            DateTime postingDate = Convert.ToDateTime(transaction.transactionDate);
            decimal transactionAmount = Convert.ToDecimal(transaction.transactionAmount);
            decimal accountBalance = Convert.ToDecimal(transaction.newBalance);

            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(postingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = string.Empty;


                description = transaction.description;
               

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);


                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (transactionAmount >= 0)
            {
                AddAccountTransactionAmount(transactionAmount, ref table); // Additions
                AddAccountTransactionAmount(0, ref table); // Subtractions
            }
            else
            {
                AddAccountTransactionAmount(0, ref table); // Additions
                AddAccountTransactionAmount(transactionAmount, ref table); // Subtractions
            }

            AddAccountBalance(accountBalance, ref table);
        }

        static void AddAccountBalance(decimal balance, ref PdfPTable table)
        {
            string amountFormatted = FormatAmount(balance);

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddAccountTransactionTitle(string title, int alignment, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddAccountTransactionAmount(decimal amount, ref PdfPTable table)
        {
            string amountFormatted = string.Empty;

            if (amount != 0)
            {
                amountFormatted = FormatAmount(amount);
            }

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddLoanTransactions(Loan loan)
        {
            Transaction[] loanTransactions;
            decimal totalFeesPeriod = 0;

            PdfPTable table = new PdfPTable(7);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 228, 48, 48, 48, 48, 54 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Eff\nDate", Element.ALIGN_LEFT, 2, ref table);
            AddLoanTransactionTitle("Transaction Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle("Interest\nCharged", Element.ALIGN_RIGHT, 2, ref table);
            AddLoanTransactionTitle("Late\nFees", Element.ALIGN_RIGHT, 2, ref table);
            AddLoanTransactionTitle("Principal", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle("Balance\nSubject to\nInterest\nRate **", Element.ALIGN_RIGHT, 4, ref table);

            if (loan.transaction != null)
            {
                loanTransactions = loan.transaction;
                foreach (Transaction t in loanTransactions)
                {
                    if (t.category.Value == "Comment")
                    {
                        AddCommentOnlyTransaction(t, ref table);
                    }
                    if(t.category.Value == "Advance")
                    {
                        AddAdvances(loanTransactions);
                    }
                    if(t.category.Value == "Payment")
                    {
                        AddLoanPaymentsSortTable(loanTransactions);
                    }
                    else
                    {
                        AddLoanTransaction(t, ref table);
                    }
                }
            }
            else
            {
                AddNoTransactionsThisPeriodMessage(ref table);
            }

            if (loan.closeDate != null)
            {
                AddLoanClosed(loan, ref table);
            }
            else
            {
                AddEndingBalance(null, loan, ref table);
            }
            if(loan.activitySummary != null)
            {
                if(loan.activitySummary.totalFeesCharged != null)
                {
                    totalFeesPeriod = Convert.ToDecimal(loan.activitySummary.totalFeesCharged);
                }
            }

            if (totalFeesPeriod > 0)
            {
                AddSeeFeeSummaryMessage(loan, ref table);
            }


            Doc.Add(table);
        }

        static void AddSeeFeeSummaryMessage(Loan loan, ref PdfPTable table)
        {
            DateTime closingDate = Convert.ToDateTime(loan.endingStatementDate);

            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(closingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("See Fee Summary Below", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }


        static void AddNoTransactionsThisPeriodMessage(ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("No Transactions This Period", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }


        static void AddLoanTransaction(Transaction transaction, ref PdfPTable table)
        {
            DateTime postingDate = Convert.ToDateTime(transaction.postingDate);
            decimal transactionAmount = 0;
            decimal interestCharged = 0;
            decimal lateFees = 0;
            decimal principal = 0;
            decimal endingBalance = 0;

            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(postingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = string.Empty;

                if(transaction.description != null)
                {
                    description = transaction.description;
                }
                

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if(transaction.grossAmount != null)
            {
                transactionAmount = Convert.ToDecimal(transaction.grossAmount);
            }
            if(transaction.interest != null)
            {
                interestCharged = Convert.ToDecimal(transaction.interest);
            }
            if(transaction.lateFee != null)
            {
                lateFees = Convert.ToDecimal(transaction.lateFee);
            }
            if(transaction.principal != null)
            {
                principal = Convert.ToDecimal(transaction.principal);
            }
            if(transaction.newBalance != null)
            {
                endingBalance = Convert.ToDecimal(transaction.newBalance);
            }

            if (transactionAmount >= 0)
            {
                AddAccountTransactionAmount(transactionAmount, ref table); // Additions
                AddAccountTransactionAmount(0, ref table); // Subtractions
            }
            else
            {
                AddAccountTransactionAmount(0, ref table); // Additions
                AddAccountTransactionAmount(transactionAmount, ref table); // Subtractions
            }

            AddLoanAccountTransactionAmount(transactionAmount, ref table); // Amount
            AddLoanAccountTransactionAmount(interestCharged, ref table); // Interest Charged
            AddLoanAccountTransactionAmount(lateFees, ref table); // Late Fees
            AddLoanAccountTransactionAmount(principal, ref table); // Principal
            AddLoanAccountTransactionAmount(endingBalance, ref table); // Balance subject to interest rate **


            AddAccountBalance(endingBalance, ref table);
        }

        static void AddAdvances(Transaction[] transactions)
        {
            int advanceCount = 0;
            DateTime postingDate = DateTime.Now;
            decimal advanceAmount = 0;
            decimal totalAdvances = 0;
            int index = 0;
            string description = string.Empty;
       
            for (int i = 0; i < transactions.Length; i++)
            {
                if (transactions[i].category.Value == "Advance")
                {
                    advanceCount++;
                }
                if (transactions[i].grossAmount != null)
                {
                    totalAdvances += Convert.ToDecimal(transactions[i].grossAmount);
                }
            }

            int rowBreakPointIndex = (int)Math.Ceiling((double)advanceCount / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("LOAN ADVANCES AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int j = 0; j < transactions.Length; j++)
            {
                while(index < advanceCount)
                {
                    if(transactions[j].category.Value == "Advance")
                    {
                        if ((index + 1) <= rowBreakPointIndex)
                        {
                            postingDate = Convert.ToDateTime(transactions[j].postingDate);
                            advanceAmount = Convert.ToDecimal(transactions[j].grossAmount);
                            if(transactions[j].description != null)
                            {
                                description = transactions[j].description;
                            }

                            rows.Add(new SortTableRow());
                            rows[index].Column.Add(postingDate.ToString("MMM dd"));
                            rows[index].Column.Add(FormatAmount(advanceAmount));
                            rows[index].Column.Add(description);
                            rows[index].Column.Add(string.Empty);
                            rows[index].Column.Add(string.Empty);
                            rows[index].Column.Add(string.Empty);
                        }
                        else
                        {
                            rows[index - rowBreakPointIndex].Column[3] = postingDate.ToString("MMM dd");
                            rows[index - rowBreakPointIndex].Column[4] = FormatAmount(advanceAmount);
                            rows[index - rowBreakPointIndex].Column[5] = description;
                        }
                    }
                    index++;
                }
                
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (advanceCount > 1)
            {
                AddSortTableSubtotal(advanceCount.ToString() + " Advances and Other Charges for " + FormatAmount(totalAdvances));
            }
        }

        static void AddLoanPaymentsSortTable(Transaction[] transactions)
        {
            int paymentCount = 0;
            DateTime postingDate = DateTime.Now;
            decimal paymentAmount = 0;
            decimal totalPayments = 0;
            int index = 0;
            string description = string.Empty;
           

            for (int i = 0; i < transactions.Length; i++)
            {
                if (transactions[i].category.Value == "Payment")
                {
                    paymentCount++;
                }
                if (transactions[i].grossAmount != null)
                {
                    totalPayments += Convert.ToDecimal(transactions[i].grossAmount);
                }
            }

            int rowBreakPointIndex = (int)Math.Ceiling((double)paymentCount / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("LOAN PAYMENTS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int j = 0; j < transactions.Length; j++)
            {
                while(index < paymentCount)
                {
                    if(transactions[j].category.Value == "Payment")
                    {
                        if ((index + 1) <= rowBreakPointIndex)
                        {
                            postingDate = Convert.ToDateTime(transactions[j].postingDate);
                            paymentAmount = Convert.ToDecimal(transactions[j].grossAmount);

                            rows.Add(new SortTableRow());
                            rows[index].Column.Add(postingDate.ToString("MMM dd"));
                            rows[index].Column.Add(FormatAmount(Math.Abs(paymentAmount)));
                            if(transactions[j].description != null)
                            {
                                description = transactions[j].description;
                            }
                            rows[index].Column.Add(description);
                            rows[index].Column.Add(string.Empty);
                            rows[index].Column.Add(string.Empty);
                            rows[index].Column.Add(string.Empty);
                        }
                        else
                        {
                            rows[index - rowBreakPointIndex].Column[3] = postingDate.ToString("MMM dd");
                            rows[index - rowBreakPointIndex].Column[4] = FormatAmount(Math.Abs(paymentAmount));
                            rows[index - rowBreakPointIndex].Column[5] = description;
                        }
                    }
                    index++;
                }
                
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (paymentCount > 1)
            {
                AddSortTableSubtotal(paymentCount.ToString() + " Payments and Other Credits for " + FormatAmount(Math.Abs(totalPayments)));
            }
        }

        static void AddDeposits(Transaction[] transactions)
        {
            int depositCount = 0;
            DateTime postingDate = DateTime.Now;
            decimal depositAmount = 0;
            decimal totalDeposits = 0;
            int index = 0;
            string description = string.Empty;

            for (int i = 0; i < transactions.Length; i++)
            {
                if (transactions[i].category.Value == "Deposit")
                {
                    depositCount++;
                }
                if (transactions[i].grossAmount != null)
                {
                    totalDeposits += Convert.ToDecimal(transactions[i].grossAmount);
                }
            }

            int rowBreakPointIndex = (int)Math.Ceiling((double)depositCount / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("DEPOSITS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

                for (int j = 0; j < transactions.Length;j++ )
                {
                        while (index < depositCount)
                        {
                            if (transactions[j].category.Value == "Deposit")
                            {
                                if ((index + 1) <= rowBreakPointIndex)
                                {
                                    postingDate = Convert.ToDateTime(transactions[j].postingDate);
                                    depositAmount = Convert.ToDecimal(transactions[j].grossAmount);
                                    

                                    rows.Add(new SortTableRow());
                                    rows[index].Column.Add(postingDate.ToString("MMM dd"));
                                    rows[index].Column.Add(FormatAmount(depositAmount));
                                    if (transactions[j].description != null)
                                    {
                                        description = transactions[j].description;
                                        rows[index].Column.Add(description);
                                    }
                                    else
                                    {
                                        rows[index].Column.Add(transactions[j].source.Value);
                                    }
                                    rows[index].Column.Add(string.Empty);
                                    rows[index].Column.Add(string.Empty);
                                    rows[index].Column.Add(string.Empty);
                                }
                                else
                                {
                                    rows[index - rowBreakPointIndex].Column[3] = postingDate.ToString("MMM dd");
                                    rows[index - rowBreakPointIndex].Column[4] = FormatAmount(depositAmount);
                                    rows[index - rowBreakPointIndex].Column[5] = description;
                                }
                               
                            }
                            index++;
                        }
                }
   
            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (depositCount > 1)
            {
                AddSortTableSubtotal(depositCount.ToString() + " Deposits and Other Credits for " + FormatAmount(totalDeposits));
            }
        }

        static void AddWithdrawals(Transaction[] transactions)
        {
            int withdrawalCount = 0;
            DateTime postingDate = DateTime.Now;
            decimal withdrawalAmount = 0;
            decimal totalWithdrawals = 0;
            int index = 0;

            for (int i = 0; i < transactions.Length; i++)
            {
                if (transactions[i].category.Value == "Withdrawal")
                {
                    withdrawalCount++;
                }
                if (transactions[i].grossAmount != null)
                {
                    totalWithdrawals += Convert.ToDecimal(transactions[i].grossAmount);
                }
            }

            int rowBreakPointIndex = (int)Math.Ceiling((double)withdrawalCount / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("WITHDRAWALS AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int j = 0; j < transactions.Length; j++)
            {
                while (index < withdrawalCount)
                {
                    if(transactions[j].category.Value == "Withdrawal")
                    {
                        if ((index + 1) <= rowBreakPointIndex)
                        {
                                postingDate = Convert.ToDateTime(transactions[j].postingDate);
                                withdrawalAmount = Convert.ToDecimal(transactions[j].grossAmount);

                                rows.Add(new SortTableRow());
                                rows[index].Column.Add(postingDate.ToString("MMM dd"));
                                rows[index].Column.Add(FormatAmount(withdrawalAmount));
                                rows[index].Column.Add(transactions[j].source.Value);
                                rows[index].Column.Add(string.Empty);
                                rows[index].Column.Add(string.Empty);
                                rows[index].Column.Add(string.Empty);
                         }
                        else
                         {
                                rows[index - rowBreakPointIndex].Column[3] = postingDate.ToString("MMM dd");
                                rows[index - rowBreakPointIndex].Column[4] = FormatAmount(withdrawalAmount);
                                rows[index - rowBreakPointIndex].Column[5] = transactions[j].source.Value;
                         }
                            
                    }
                    index++;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (withdrawalCount > 1)
            {
                AddSortTableSubtotal(withdrawalCount.ToString() + " Withdrawals and Other Charges for " + FormatAmount(totalWithdrawals));
            }
        }

        static void AddYearToDateTotals(Loan loan)
        {
            decimal totalFeesCharged = 0;
            decimal totalInterestCharged = 0;
            decimal totalLoanFeesLastYear = 0;
            decimal totalInterestLastYear = 0;
            DateTime openDate = Convert.ToDateTime(loan.openDate);

            if(loan.loanFeesChargedYTD != null)
            {
                totalFeesCharged = Convert.ToDecimal(loan.loanFeesChargedYTD);
            }
            if(loan.interestChargedYTD != null)
            {
                totalInterestCharged = Convert.ToDecimal(loan.interestChargedYTD);
            }
            if(loan.loanFeesChargedLastYear != null)
            {
                totalLoanFeesLastYear = Convert.ToDecimal(loan.loanFeesChargedLastYear);
            }
            if(loan.interestChargedLastYear != null)
            {
                totalInterestLastYear = Convert.ToDecimal(loan.interestChargedLastYear);
            }

   

            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 12, 153, 79, 93, 188 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 20f;
            table.KeepTogether = true;

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("YEAR TO DATE TOTALS", GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Fees Charged this Year Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Fees Charged this Year", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // Total Fees Charged this Year Value
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(FormatAmount(totalFeesCharged), GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingRight = 35f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Interest Charged this Year Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Interest Charged this Year", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingBottom = (true) ? 0f : 5f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = (true) ? 0f : 1f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // Total Interest Charged this Year Value
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(FormatAmount(totalInterestCharged), GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingBottom = (true) ? 0f : 5f;
                cell.PaddingRight = 35f;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = (true) ? 0f : 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingBottom = (true) ? 0f : 5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = (true) ? 0f : 1f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (openDate < DateTime.Now)
            {
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Total Fees Charged Last Year
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Total Fees Charged Last Year", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingLeft = 6f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderWidthLeft = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // Total Interest Charged this Year Value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(FormatAmount(totalLoanFeesLastYear), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingRight = 35f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderWidthRight = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Total Interest Charged Last Year
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Total Interest Charged Last Year", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingLeft = 6f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderWidthLeft = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // Total Interest Charged Last Year Value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(FormatAmount(totalInterestLastYear), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.PaddingRight = 35f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderWidthRight = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
            }



            Doc.Add(table);
        }


        static void AddFeeSummary(Loan loan)
        {
            decimal totalFees = 0;

            if(loan.activitySummary != null)
            {
                if (loan.activitySummary.totalFeesCharged != null)
                {
                    totalFees = Convert.ToDecimal(loan.activitySummary.totalFeesCharged);
                }
            }

            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Title
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 12f;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("FEE SUMMARY", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 15;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 10f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            AddLoanFees(loan);

            // TOTAL FEES FOR THIS PERIOD
            {
                PdfPTable table = new PdfPTable(4);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 51, 234, 48, 192 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // TOTAL FEES FOR THIS PERIOD
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("TOTAL FEES FOR THIS PERIOD", GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_LEFT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(totalFees.ToString("N"), GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }

        

        static void AddLoanFees(Loan loan)
        {
            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);

            Transaction[] fees;
            if(loan.feeTransaction != null)
            {
                fees = loan.feeTransaction;

                foreach(Transaction fee in fees)
                {
                    AddLoanFee(fee, ref table);
                }
            }

            Doc.Add(table);
        }

        static void AddInterestChargedSummary(Loan loan)
        {
            decimal totalInterest = 0;
            InterestCharge[] interestCharges;
            decimal newBalance = 0;

            if(loan.activitySummary != null)
            {
                if(loan.activitySummary.totalInterestCharged != null)
                {
                    totalInterest = Convert.ToDecimal(loan.activitySummary.totalInterestCharged);
                }
            }

            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Title
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 12f;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("INTEREST CHARGED SUMMARY", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                //p.IndentationLeft = 15;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 10f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            if(loan.ending.balance != null)
            {
                newBalance = Convert.ToDecimal(loan.ending.balance);
            }
            if(loan.interestCharge != null)
            {
                interestCharges = loan.interestCharge;
                AddLoanInterestTransactions(interestCharges, newBalance);
            }
            

            // TOTAL INTEREST FOR THIS PERIOD
            {
                PdfPTable table = new PdfPTable(4);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 51, 234, 48, 192 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // TOTAL INTEREST FOR THIS PERIOD
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("TOTAL INTEREST FOR THIS PERIOD", GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_LEFT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(totalInterest.ToString("N"), GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }

        static void AddLoanInterestTransactions(InterestCharge[] charges, decimal newBal)
        {
            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);

            foreach (InterestCharge interest in charges)
            {
                decimal interestAmount = Convert.ToDecimal(interest.interest);

                if (interestAmount > 0)
                {
                    AddLoanInterestTransaction(interest, newBal, ref table);
                }
            }

            Doc.Add(table);
        }

        static void AddLoanInterestTransaction(InterestCharge charge, decimal newBal, ref PdfPTable table)
        {
            DateTime postingDate = DateTime.Now;
            decimal interestAmount = 0;
            decimal newBalance = newBal;
            string description = string.Empty;

            if(charge.postingDate != null)
            {
                postingDate = Convert.ToDateTime(charge.postingDate);
            }
            if(charge.interest != null)
            {
                interestAmount = Convert.ToDecimal(charge.interest);
            }
            

            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(postingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddLoanAccountTransactionAmount(interestAmount, ref table); // Amount

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountBalance(newBalance, ref table);
        }


        static void AddLoanFee(Transaction fee, ref PdfPTable table)
        {
            DateTime feePostingDate = DateTime.Now;
            string description = string.Empty;
            decimal feeAmount = 0;
            decimal newBalance = 0;

            if(fee.postingDate != null)
            {
                feePostingDate = Convert.ToDateTime(fee.postingDate);
            }
            if(fee.description != null)
            {
                description = fee.description;
            }
            if(fee.grossAmount != null)
            {
                feeAmount = Convert.ToDecimal(fee.grossAmount);
            }
            if(fee.newBalance != null)
            {
                newBalance = Convert.ToDecimal(fee.newBalance);
            }

            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(feePostingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddLoanAccountTransactionAmount(feeAmount, ref table); // Amount

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountBalance(newBalance, ref table);
        }


        static void AddLoanAccountTransactionAmount(decimal amount, ref PdfPTable table)
        {
            string amountFormatted = FormatAmount(amount);

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetNormalFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddLoanTransactionTitle(string title, int alignment, int numOfLines, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(10f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.GRAY;
            cell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            cell.AddElement(Underline(chunk, alignment, numOfLines));
            table.AddCell(cell);
        }


        static void AddCommentOnlyTransaction(Transaction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = string.Empty;
                if(transaction.description != null)
                {
                    description = transaction.description;
                }
                
                

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 20;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.NoWrap = true;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table); // Additions
            AddAccountTransactionAmount(0, ref table); // Subtractions
            AddAccountTransactionAmount(0, ref table);
        }


        static void AddLoanClosed(Loan loan, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(loan.description + " " + "Closed\n*** This is the final statement you will receive for this account ***\n*** Please retain this final statement for tax reporting purposes ***", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.NoWrap = true;
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        static void AddShareClosed(Share closedShare, ref PdfPTable table)
        {
            // Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }

            // Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(closedShare.description + " " + "Closed\n*** This is the final statement you will receive for this account ***\n*** Please retain this final statement for tax reporting purposes ***", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.NoWrap = true;
                cell.PaddingTop = -1f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
        }

        static void AddEndingBalance(Share share, Loan loan, ref PdfPTable table)
        {

            DateTime endDate = DateTime.Now; 
            decimal endingBalance = 0;
            
            if(share == null)
            {
                endDate = Convert.ToDateTime(loan.endingStatementDate);
                endingBalance = Convert.ToDecimal(loan.ending.balance);
            }
            else if(loan == null)
            {
                endDate = Convert.ToDateTime(share.endingStatementDate);
                endingBalance = Convert.ToDecimal(share.ending.balance);
            }
          

            // Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(endDate.ToString("MMM dd"), GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }

            // Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Ending Balance", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountBalance(endingBalance, ref table);
        }

        static void AddTotalFees(Share share)
        {
            decimal totalOverdraftFeesPeriod = 0;
            decimal totalOverdraftFeesYTD = 0;
            decimal totalReturnedItemFeePeriod = 0;
            decimal totalReturnedItemFeeYTD = 0;

            if(share.overdraftFeePeriod != null)
            {
                totalOverdraftFeesPeriod = Convert.ToDecimal(share.overdraftFeePeriod);
            }
            if(share.overdraftFeeYTD != null)
            {
                totalOverdraftFeesYTD = Convert.ToDecimal(share.overdraftFeeYTD);
            }
            if(share.returnedItemFeePeriod != null)
            {
                totalReturnedItemFeePeriod = Convert.ToDecimal(share.returnedItemFeePeriod);
            }
            if(share.returnedItemFeeYTD != null)
            {
                totalReturnedItemFeeYTD = Convert.ToDecimal(share.returnedItemFeeYTD);
            }



            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 12, 153, 79, 93, 188 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddFeeTitle("", new Border(1, 1, 0, 1), ref table);
            AddFeeTitle("Total for\nthis period", new Border(1, 1, 0, 1), ref table);
            AddFeeTitle("Total\nyear-to-date", new Border(1, 1, 0, 1), ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Overdraft Fees
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Overdraft Fees", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }
            AddFeeValue(totalOverdraftFeesPeriod, new Border(1, 0, 0, 0), -2f, ref table);
            AddFeeValue(totalOverdraftFeesYTD, new Border(1, 0, 0, 0), -2f, ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Returned Item Fees
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Returned Item Fees", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 1;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            AddFeeValue(totalReturnedItemFeePeriod, new Border(1, 0, 0, 1), -8f, ref table);
            AddFeeValue(totalReturnedItemFeeYTD, new Border(1, 0, 0, 1), -8f, ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            Doc.Add(table);
        }

        static void AddFeeTitle(string title, Border border, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(11f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            cell.AddElement(p);
            cell.Padding = 6f;
            cell.PaddingTop = -2f;
            cell.BorderWidth = 0f;
            cell.BorderWidthLeft = border.WidthLeft;
            cell.BorderWidthTop = border.WidthTop;
            cell.BorderWidthRight = border.WidthRight;
            cell.BorderWidthBottom = border.WidthBottom;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
        }

        static void AddFeeValue(decimal value, Border border, float paddingTop, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(FormatAmount(value), GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.Padding = 6f;
            cell.PaddingTop = paddingTop;
            cell.PaddingRight = 35f;
            cell.BorderWidth = 0f;
            cell.BorderWidthLeft = border.WidthLeft;
            cell.BorderWidthTop = border.WidthTop;
            cell.BorderWidthRight = border.WidthRight;
            cell.BorderWidthBottom = border.WidthBottom;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
        }


        static void AddSortTableHeading(string title)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            cell.AddElement(p);
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddSortTableTitle(string title, int alignment, ref PdfPTable table)
        {
            AddSortTableTitle(title, alignment, 0, ref table);
        }

        static void AddSortTableTitle(string title, int alignment, float indentation, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            p.IndentationLeft = indentation;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddSortTableValue(string value, int alignment, ref PdfPTable table)
        {
            AddSortTableValue(value, alignment, 0, ref table);
        }

        static void AddSortTableValue(string value, int alignment, float indentation, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetNormalFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            p.IndentationLeft = indentation;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddSortTableSubtotal(string value)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetBoldItalicFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 70;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddMoneyPerksSummary(MoneyPerksStatement moneyPerksStatement)
        {
            if (moneyPerksStatement != null)
            {
                AddSectionHeading("MONEYPERKS POINTS SUMMARY");

                PdfPTable table = new PdfPTable(5);
                table.HeaderRows = 1;
                float[] tableWidths = new float[] { 51, 280, 62, 65, 67 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                AddMoneyPerksTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
                AddMoneyPerksTransactionTitle("Transaction Description", Element.ALIGN_LEFT, 1, ref table);
                AddMoneyPerksTransactionTitle("Points\nAwarded", Element.ALIGN_RIGHT, 2, ref table);
                AddMoneyPerksTransactionTitle("Points\nRedeemed", Element.ALIGN_RIGHT, 2, ref table);
                AddMoneyPerksTransactionTitle("Balance", Element.ALIGN_RIGHT, 1, ref table);
                AddMoneyPerksBeginningBalance(moneyPerksStatement, ref table);

                foreach (MoneyPerksTransaction transaction in moneyPerksStatement.Transactions)
                {
                    AddMoneyPerksTransaction(transaction, ref table);
                }

                AddMoneyPerksEndingBalance(moneyPerksStatement, ref table);
                Doc.Add(table);
            }
        }

        static void AddMoneyPerksTransactionTitle(string title, int alignment, int numOfLines, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(10f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.GRAY;
            cell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            cell.AddElement(Underline(chunk, alignment, numOfLines));
            table.AddCell(cell);
        }

        static void AddMoneyPerksBalance(int balance, ref PdfPTable table)
        {
            string amountFormatted = balance.ToString();

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddMoneyPerksBeginningBalance(MoneyPerksStatement moneyPerksStatement, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Beginning Balance", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksBalance(moneyPerksStatement.BeginningBalance, ref table);
        }

        static void AddMoneyPerksEndingBalance(MoneyPerksStatement moneyPerksStatement, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Ending Balance", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksBalance(moneyPerksStatement.EndingBalance, ref table);
        }

        static void AddMoneyPerksTransaction(MoneyPerksTransaction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(transaction.Date.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = transaction.Description;

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (transaction.Amount >= 0)
            {
                AddMoneyPerksTransactionAmount(transaction.Amount, ref table); // Additions
                AddMoneyPerksTransactionAmount(0, ref table); // Subtractions
            }
            else
            {
                AddMoneyPerksTransactionAmount(0, ref table); // Additions
                AddMoneyPerksTransactionAmount(transaction.Amount, ref table); // Subtractions
            }

            AddMoneyPerksBalance(transaction.Balance, ref table);
        }

        static void AddMoneyPerksTransactionAmount(int amount, ref PdfPTable table)
        {
            string amountFormatted = string.Empty;

            if (amount != 0)
            {
                amountFormatted = amount.ToString();
            }

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddAccountSubHeading(string subtitle, bool stroke)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            float cellPaddingTop = -1f;

            if (stroke)
            {
                cellPaddingTop = -6f;
                AddSubHeadingStroke();
            }

            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(subtitle, GetBoldFont(12f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            cell.AddElement(p);
            cell.PaddingTop = cellPaddingTop;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddLoanTransactionsFooter(string value)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 12f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(11f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 15;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddBalanceForward(Share share, ref PdfPTable table)
        {
            DateTime beginDate = Convert.ToDateTime(share.beginningStatementDate);
            decimal balanceForward = Convert.ToDecimal(share.beginning.balance);

            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(beginDate.ToString("MMM dd"), GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Balance Forward", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountBalance(balanceForward, ref table);
        }

        static void AddHeaderBalances(StatementProduction statement, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk("Account  Balances  at  a  Glance:", GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            decimal testValue = 250;
            p.IndentationLeft = 81;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            PdfPTable balancesTable = new PdfPTable(2);
            float[] tableWidths = new float[] { 60f, 40f };
            balancesTable.SetWidthPercentage(tableWidths, Doc.PageSize);
            balancesTable.TotalWidth = 300f;
            balancesTable.LockedWidth = true;
            AddHeaderBalanceTitle("Total  Checking:", ref balancesTable);
            AddHeaderBalanceValue(testValue, ref balancesTable);
            AddHeaderBalanceTitle("Total  Savings:", ref balancesTable);
            AddHeaderBalanceValue(testValue, ref balancesTable);
            AddHeaderBalanceTitle("Total  Loans:", ref balancesTable);
            AddHeaderBalanceValue(testValue, ref balancesTable);
            AddHeaderBalanceTitle("Total  Certificates:", ref balancesTable);
            AddHeaderBalanceValue(testValue, ref balancesTable);
            cell.AddElement(balancesTable);
            table.AddCell(cell);
        }

        static void AddHeaderBalanceTitle(string title, ref PdfPTable balancesTable)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetNormalFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 75;
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            balancesTable.AddCell(cell);
        }

        static void AddHeaderBalanceValue(decimal value, ref PdfPTable balancesTable)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(FormatAmount(value), GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            p.IndentationRight = 28;
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            balancesTable.AddCell(cell);
        }

        static void AddInvisibleAccountNumber(StatementProduction statement)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 612f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(statement.envelope[0].statement[0].account[0].accountNumber, GetNormalFont(5f));
            Paragraph p = new Paragraph(chunk);
            p.Font.SetColor(255, 255, 255);
            cell.AddElement(p);
            cell.PaddingTop = -9f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddStatementHeading(string text, float indentationLeft, float paddingTop)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 612f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(text, GetNormalFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = indentationLeft;
            cell.AddElement(p);
            cell.PaddingTop = paddingTop;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddTopAdvertising(StatementProduction statement)
        {
            // Advertisement Bottom
            Font font = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD, new BaseColor(0, 0, 0));
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 34f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = null;

        /*
            
            for (int i = 0; i < statement.AdvertisementTop.TotalLines; i++)
            {
                if (chunk == null)
                {
                    chunk = new Chunk(statement.AdvertisementTop.MessageLines[i], font);
                }
                else
                {
                    chunk.Append("\n" + statement.AdvertisementTop.MessageLines[i]);
                }
            }
            */
            if (chunk == null)
            {
                chunk = new Chunk(string.Empty, font);
                table.SpacingBefore = 14f;
            }

            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(12f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            p.IndentationLeft = 385;
            cell.AddElement(p);
            cell.PaddingTop = -1f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);

            Doc.Add(table);
        }

        static void AddBottomAdvertising(StatementProduction statement)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Advertisement Bottom Stroke
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            PdfPCell cell = new PdfPCell();
            cell.BorderWidthBottom = 5f;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
            Doc.Add(table);

            // Advertisement Bottom
            Font font = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD, new BaseColor(0, 0, 0));
            table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            cell = new PdfPCell();
            Chunk chunk = null;
            
            for (int i = 0; i < statement.AdvertisementBottom.TotalLines; i++)
            {
                if (chunk == null)
                {
                    chunk = new Chunk(statement.AdvertisementBottom.MessageLines[i], font);
                }
                else
                {
                    chunk.Append("\n" + statement.AdvertisementBottom.MessageLines[i]);
                }
            }
            
            if (chunk != null)
            {
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(12f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_CENTER;
                p.IndentationLeft = 385;
                cell.AddElement(p);
                cell.PaddingTop = -1f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                Doc.Add(table);
            }
        }


        static void AddPageNumbersAndDisclosures(StatementProduction statement)
        {
            string beginDate = statement.envelope[0].statement[0].beginningStatementDate;
            string endDate = statement.envelope[0].statement[0].endingStatementDate;
            string firstAccNumber = statement.envelope[0].statement[0].account[0].accountNumber;

            // Adds page numbers
            PdfReader statementReader = new PdfReader("C:\\" + TEMP_FILE_NAME);
            PdfReader statementBackReader = new PdfReader(Configuration.GetStatementDisclosuresTemplateFilePath());

            using (FileStream fs = new FileStream(Configuration.GetStatementsOutputPath() + firstAccNumber + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (PdfStamper stamper = new PdfStamper(statementReader, fs))
                {
                    stamper.SetFullCompression();
                    int pageCount = statementReader.NumberOfPages + 1; // Adds 1 for the disclosures page that will be added later
                    for (int i = 1; i <= pageCount - 1; i++)
                    {
                        if (i == 1)
                        {
                            // Page count on first page
                            Chunk chunk = new Chunk("Page:   1 of   " + pageCount, GetBoldFont(12f));
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(chunk), 578, 595, 0);
                            if (i != pageCount)
                            {
                                chunk = new Chunk("--- Continued on following page ---", GetBoldFont(9f));
                                ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase(chunk), 300, 20, 0);
                            }
                        }
                        else if (i != pageCount)
                        {
                            float startY = 750f;
                            float lineHeight = 10;

                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk(beginDate + "  thru  " + endDate, GetBoldFont(9f))), 578, startY, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Account  Number:   ******" + firstAccNumber.ToString(), GetBoldFont(9f))), 578, startY - lineHeight, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Page:  " + i.ToString() + " of " + pageCount.ToString(), GetBoldFont(9f))), 578, startY - (lineHeight * 2), 0);
                            Chunk chunk = new Chunk("--- Continued on reverse side ---", GetBoldFont(9f));
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase(chunk), 300, 20, 0);
                        }
                        else
                        {
                            float startY = 750f;
                            float lineHeight = 10;
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk(beginDate + "  thru  " + endDate, GetBoldFont(9f))), 578, startY, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Account  Number:   ******" + firstAccNumber.ToString(), GetBoldFont(9f))), 578, startY - lineHeight, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Page:  " + i.ToString() + " of " + pageCount.ToString(), GetBoldFont(9f))), 578, startY - (lineHeight * 2), 0);
                        }
                    }

                    stamper.InsertPage(pageCount, PageSize.LETTER);
                    PdfContentByte cb = stamper.GetOverContent(pageCount);
                    PdfImportedPage p = stamper.GetImportedPage(statementBackReader, 1);
                    cb.AddTemplate(p, 0, 0);
                }
            }
        }

        static void AddLoanPaymentInformation(Loan loan)
        {
            decimal creditLimit = 0;
            decimal annualPercentageRate = 0;
            decimal availableCredit = 0;
            decimal previousBalance = 0;
            decimal newBalance = 0;
            decimal minimumPaymentDue = 0;
            DateTime paymentDueDate = DateTime.Now;
            DateTime nextPaymentDueDate = DateTime.Now;

            if(loan.creditLimit != null)
            {
                creditLimit = Convert.ToDecimal(loan.creditLimit);
            }
            if(loan.creditLimitAvailable != null)
            {
                availableCredit = Convert.ToDecimal(loan.creditLimitAvailable);
            }
            if(loan.beginning.annualRate != null)
            {
                annualPercentageRate = Convert.ToDecimal(loan.beginning.annualRate);
            }
            if(loan.beginning.balance != null)
            {
                previousBalance = Convert.ToDecimal(loan.beginning.balance);
            }
            if(loan.ending.balance != null)
            {
                newBalance = Convert.ToDecimal(loan.ending.balance);
            }
            if(loan.endingDueDate != null)
            {
                paymentDueDate = Convert.ToDateTime(loan.endingDueDate);
            }
            if(loan.nextScheduledDueDate != null)
            {
                nextPaymentDueDate = Convert.ToDateTime(loan.nextScheduledDueDate);
            }

            if(loan.paymentSummary != null)
            {
                if(loan.paymentSummary.minimumPaymentDueAmount != null)
                {
                    minimumPaymentDue = Convert.ToDecimal(loan.paymentSummary.minimumPaymentDueAmount);
                }
            }


            // Annual percentage rate
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = null;
                if (creditLimit <= 0)
                {
                    chunk = new Chunk("Annual Percentage Rate:  " + annualPercentageRate.ToString("N3") + "%", GetBoldFont(9f));
                }
                else
                {
                    chunk = new Chunk("Annual Percentage Rate:  " + annualPercentageRate.ToString("N3") + "%    Credit Limit:    " + FormatAmount(creditLimit) + "    Available Credit:    " + FormatAmount(availableCredit), GetBoldFont(9f));
                }
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 28f;
                cell.AddElement(p);
                cell.PaddingTop = 7f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            // PAYMENT INFORMATION
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("PAYMENT INFORMATION", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = 7f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            // Summary table
            {
                PdfPTable table = new PdfPTable(3);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 93, 55, 381 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // Previous Balance Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Previous Balance:", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Previous Balance
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(previousBalance.ToString("N"), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // New Balance Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("New Balance:", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // New Balance
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(newBalance.ToString("N"), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Minimum Payment Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Minimum Payment:", GetBoldFont(9f));

                    if (minimumPaymentDue == 0)
                    {
                        chunk = new Chunk("Minimum Payment: No Payment Due", GetBoldFont(9f));
                    }
                   
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Minimum Payment
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    if(minimumPaymentDue != 0)
                    {
                        chunk = new Chunk(minimumPaymentDue.ToString("N"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Payment Due Date Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Payment Due Date:", GetBoldFont(9f));
                    if (paymentDueDate.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        chunk = new Chunk("Payment Due Date: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Payment Due Date
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    if (paymentDueDate.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        chunk = new Chunk(paymentDueDate.ToString("MM/dd/yyyy"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }

            // Next Payment Due Date after statement
            {
                PdfPTable table = new PdfPTable(3);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 180, 55, 290 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0;

                // Next Payment Due Date after statement Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Next Payment Due Date after statement:", GetBoldFont(9f));
                    if (nextPaymentDueDate.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        chunk = new Chunk("Next Payment Due Date after statement: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Next Payment Due Date after statement
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    if (nextPaymentDueDate.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        chunk = new Chunk(nextPaymentDueDate.ToString("MM/dd/yyyy"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }
        /*
        static void AddYtdSummaries(Share share, Loan loan)
        {
            PdfPTable leftTable;
            PdfPCell leftTableCell;

            AddSectionHeading("YTD SUMMARIES");


            // A table to create 2 columns
            {
                PdfPTable table = new PdfPTable(2);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 255, 270 };
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 15f;

                // TOTAL DIVIDENDS PAID
                {
                    leftTable = new PdfPTable(2);
                    leftTable.HeaderRows = 0;
                    float[] leftTableWidths = new float[] { 214.5f, 48f };
                    leftTable.TotalWidth = 262.5f;
                    leftTable.SetWidths(leftTableWidths);
                    leftTable.LockedWidth = true;
                    leftTable.SpacingBefore = 0f;

                    // Title
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("TOTAL DIVIDENDS PAID", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // For layout only
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                    // Account
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(statement.envelope[0].statement[0].account[0].typeDescription, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(FormatAmount(account.Dividends), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        chunk = new Chunk(irsContributionYtd.Description, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        chunk = new Chunk(FormatAmount(irsContributionYtd.Amount), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.DividendsTotalSet)
                        {
                            chunk = new Chunk("Total Year To Date Dividends Paid", GetNormalFont(9f));
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.DividendsTotalSet)
                        {
                            chunk = new Chunk(FormatAmount(statement.DividendsTotal), GetNormalFont(9f));
                        }
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.NontaxableDividendYtdSet)
                        {
                            chunk = new Chunk("Nontaxable Dividends", GetNormalFont(9f));
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        if (statement.NontaxableDividendYtdSet)
                        {
                            chunk = new Chunk(FormatAmount(statement.NontaxableDividendYtd), GetNormalFont(9f));
                        }
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        chunk = new Chunk(ytdTotal.Description, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        chunk = new Chunk(FormatAmount(ytdTotal.Amount), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.InterestPaidTotalYtdSet)
                        {
                            chunk = new Chunk("Total Year To Date Interest Paid", GetNormalFont(9f));
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        if (statement.InterestPaidTotalYtdSet)
                        {
                            chunk = new Chunk(FormatAmount(statement.InterestPaidTotalYtd), GetNormalFont(9f));
                        }
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    leftTableCell = new PdfPCell();
                    if (statement.Accounts.Count > 0) leftTableCell.AddElement(leftTable);
                    leftTableCell.BorderWidth = 0;
                    leftTableCell.Padding = 0;

                    table.AddCell(leftTableCell);
                }

                // TOTAL LOAN INTEREST PAID
                {
                    leftTable = new PdfPTable(2);
                    leftTable.HeaderRows = 0;
                    float[] leftTableWidths = new float[] { 214.5f, 48f };
                    leftTable.TotalWidth = 262.5f;
                    leftTable.SetWidths(leftTableWidths);
                    leftTable.LockedWidth = true;
                    leftTable.SpacingBefore = 0f;

                    // Title
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("TOTAL LOAN INTEREST PAID", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // For layout only
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }



                    // Loan
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(loan.Description, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(FormatAmount(loan.TotalInterestChargedYtd), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    decimal loanCount = Convert.ToDecimal(statement.epilogue.loanCount);
                    leftTableCell = new PdfPCell();
                    if (loanCount > 0) leftTableCell.AddElement(leftTable);
                    leftTableCell.BorderWidth = 0;
                    leftTableCell.Padding = 0;

                    table.AddCell(leftTableCell);
                }



                Doc.Add(table);
            }
        }
        */
        static void AddSectionHeading(string title)
        {
            if (Writer.GetVerticalPosition(false) <= 175)
            {
                Doc.NewPage();
            }

            AddHeadingStroke();
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(16f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddHeadingStroke()
        {
            Doc.Add(Stroke(525f, 20f, 0, 5f, BaseColor.BLACK, Element.ALIGN_CENTER));
        }

        static void AddSubHeadingStroke()
        {
            Doc.Add(Stroke(525f, 10f, 0, 0.5f, BaseColor.BLACK, Element.ALIGN_CENTER));
        }


        static PdfPTable Stroke(float width, float spacingAbove, float spacingLeft, float thickness, BaseColor color, int alignment)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = width;
            table.LockedWidth = true;
            table.SpacingBefore = spacingAbove;
            PdfPCell cell = new PdfPCell();
            cell.BorderWidth = 0;
            cell.BorderWidthBottom = thickness;
            cell.BorderColor = color;
            cell.PaddingLeft = spacingLeft;
            table.AddCell(cell);
            table.HorizontalAlignment = alignment;
            return table;
        }

        /// <summary>
        /// Produces a stroke with a width that will fit underneath a chunk of text, even if the text is multiple lines long
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        static PdfPTable Underline(Chunk textChunk, int alignment, int numOfLines)
        {
            if (numOfLines > 1)
            {
                string[] words = textChunk.ToString().Split('\n');
                string longestWord = string.Empty;
                Chunk wordChunk = null;
                Chunk longestWordChunk = new Chunk(longestWord);

                foreach (string word in words)
                {
                    wordChunk = new Chunk(word);
                    wordChunk.Font = textChunk.Font;
                    wordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                    longestWordChunk = new Chunk(longestWord);
                    longestWordChunk.Font = textChunk.Font;
                    longestWordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                    if (wordChunk.GetWidthPoint() > longestWordChunk.GetWidthPoint())
                    {
                        longestWord = word;
                    }
                }

                longestWordChunk = new Chunk(longestWord);
                longestWordChunk.Font = textChunk.Font;
                longestWordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                return Stroke(longestWordChunk.GetWidthPoint(), -3f, 0, 0.5f, BaseColor.BLACK, alignment);
            }
            else
            {
                return Stroke(textChunk.GetWidthPoint(), -3f, 0, 0.5f, BaseColor.BLACK, alignment);
            }
        }


        static Font GetNormalFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.NORMAL, new BaseColor(0, 0, 0));
        }

        static Font GetBoldFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.BOLD, new BaseColor(0, 0, 0));
        }

        static Font GetItalicFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.ITALIC, new BaseColor(0, 0, 0));
        }

        static Font GetBoldItalicFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.BOLDITALIC, new BaseColor(0, 0, 0));
        }

        static string FormatAmount(decimal amount)
        {
            string formattedAmount = amount.ToString("N");

            //Puts negative sign at end
            if (formattedAmount.StartsWith("-"))
            {
                formattedAmount = formattedAmount.Replace("-", string.Empty);
                formattedAmount += "-";
            }

            return formattedAmount;
        }

        public static int GetNumberOfStatementsBuilt()
        {
            return NumberOfStatementsBuilt;
        }

        static Document Doc
        {
            get;
            set;
        }

        static PdfWriter Writer
        {
            get;
            set;
        }

        private static int NumberOfStatementsBuilt
        {
            get;
            set;
        }

        static Advertisement[] AdvertisementTop
        {
            get;
            set;
        }

        public static string TEMP_FILE_NAME = "statement_pdf.temp";
        public const int MAX_RELATIONSHIP_BASED_LEVELS = 10;
    }

    class StatementPageEvent : PdfPageEventHelper
    {
        public override void OnStartPage(PdfWriter writer, Document Document)
        {
            string nextPageTemplate = Configuration.GetStatementTemplateFilePath();

            if (Document.PageNumber > 1)
            {
                Document.SetMargins(STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_TOP, STATEMENT_MARGIN_BOTTOM);
                Document.NewPage();

                using (FileStream templateInputStream = File.Open(nextPageTemplate, FileMode.Open))
                {
                    // Loads existing PDF
                    PdfReader reader = new PdfReader(templateInputStream);
                    PdfContentByte contentByte = writer.DirectContent;
                    PdfImportedPage page = writer.GetImportedPage(reader, 1);

                    // Copies first page of existing PDF into output PDF
                    //Document.NewPage();
                    contentByte.AddTemplate(page, 0, 0);
                }
            }
            else
            {
                Document.SetMargins(STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_SIDES, FIRST_PAGE_STATEMENT_MARGIN_TOP, STATEMENT_MARGIN_BOTTOM);
            }
        }


        public static float FIRST_PAGE_STATEMENT_MARGIN_TOP = 12f;
        public static float STATEMENT_MARGIN_TOP = 70f;
        public static float STATEMENT_MARGIN_BOTTOM = 30f;
        public static float STATEMENT_MARGIN_SIDES = 12f;
    }
}
