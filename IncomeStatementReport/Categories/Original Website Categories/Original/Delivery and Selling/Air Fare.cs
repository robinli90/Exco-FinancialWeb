﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcoUtility;

namespace IncomeStatementReport.Categories.Delivery_and_Selling
{
    public class Air_Fare : Group
    {
        public Air_Fare(int fiscalYear, int fiscalMonth)
        {
            name = "AIR FARE";
            // add accounts
            plant01.accountList.Add(new Account("100", "506100"));
            plant03.accountList.Add(new Account("300", "506100"));
            plant05.accountList.Add(new Account("500", "506100"));
            plant04.accountList.Add(new Account("452", "551501"));
            plant04.accountList.Add(new Account("452", "551502"));
            plant41.accountList.Add(new Account("4152", "551501"));
            plant48.accountList.Add(new Account("4852", "551501"));
            plant49.accountList.Add(new Account("4952", "551501"));

            plant41.accountList.Add(new Account("4152", "551502"));

            // process accounts
            plant01.GetAccountsData(fiscalYear, fiscalMonth);
            plant03.GetAccountsData(fiscalYear, fiscalMonth);
            plant05.GetAccountsData(fiscalYear, fiscalMonth);
            plant04.GetAccountsData(fiscalYear, fiscalMonth);
            plant41.GetAccountsData(fiscalYear, fiscalMonth);
            plant48.GetAccountsData(fiscalYear, fiscalMonth);
            plant49.GetAccountsData(fiscalYear, fiscalMonth);
        }
    }
}
