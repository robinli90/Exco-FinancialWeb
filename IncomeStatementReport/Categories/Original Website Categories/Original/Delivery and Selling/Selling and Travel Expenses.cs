﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcoUtility;

namespace IncomeStatementReport.Categories.Delivery_and_Selling
{
    public class Selling_and_Travel_Expenses : Group
    {
        public Selling_and_Travel_Expenses(int fiscalYear, int fiscalMonth)
        {
            name = "SELLING AND TRAVEL EXPENSES";
            // add accounts
            plant01.accountList.Add(new Account("100", "506000"));
            plant03.accountList.Add(new Account("300", "506000"));
            plant05.accountList.Add(new Account("500", "506000"));
            plant04.accountList.Add(new Account("452", "550501"));
            plant04.accountList.Add(new Account("452", "550502"));
            plant04.accountList.Add(new Account("452", "552001"));
            plant04.accountList.Add(new Account("452", "552002"));
            plant41.accountList.Add(new Account("4152", "550501"));
            plant41.accountList.Add(new Account("4152", "552001"));
            plant41.accountList.Add(new Account("4152", "552002"));
            plant48.accountList.Add(new Account("4852", "550501"));
            plant48.accountList.Add(new Account("4852", "550502"));
            plant48.accountList.Add(new Account("4852", "552001"));
            plant49.accountList.Add(new Account("4952", "550501"));
            plant49.accountList.Add(new Account("4952", "552001"));
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
