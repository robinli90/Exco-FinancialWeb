using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ExcoUtility;
using System.Data.Odbc;
using System.Net.Mail;
using System.Net;

namespace SalesReport
{
    // NOTE:
    // All four plants starts from Oct to September next year
    // Markham, Michigan and Texas follow fiscal period
    // Colombia breaks fiscal period

    public class Process
    {
        public Plant plant;
        // invoice map
        public Dictionary<int, Invoice> invoiceMap = new Dictionary<int, Invoice>();
        // customer bill-to map
        public Dictionary<string, ExcoCustomer> customerBillToMap = new Dictionary<string, ExcoCustomer>();
        // customer ship-to map
        public Dictionary<string, ExcoCustomer> customerShipToMap = new Dictionary<string, ExcoCustomer>();
        // budget map
        public Dictionary<string, List<ExcoMoney>> budgetMap = new Dictionary<string, List<ExcoMoney>>();
        // order and parts map
        public Dictionary<int, List<string>> partMap = new Dictionary<int, List<string>>();
        // flags of including steel surcharge
        public bool doesIncludeSurcharge = false;
        // fiscal yera
        public int fiscalYear = 16;
        // file path
        public string filePath = string.Empty;

        // constructor
        public Process()
        {
            // get list of customers
            List<ExcoCustomer> customerList = ExcoCustomerTable.Instance.GetAllCustomers();
            Console.WriteLine("Get All Customers Done");
            // fill bill-to map and ship-to map
            foreach (ExcoCustomer customer in customerList)
            {
                customerBillToMap.Add(customer.BillToID, customer);
                customerShipToMap.Add(customer.ShipToID, customer);
            }
        }

        public void Run(int plantID)
        {
            switch (plantID)
            {
                case 1:
                    plant = new Plant(plantID, "CA");
                    break;
                case 3:
                case 5:
                    plant = new Plant(plantID, "US");
                    break;
                case 4:
                    plant = new Plant(plantID, "CP");
                    break;
                default:
                    throw new Exception("Invalid plant ID: " + plantID.ToString());
            }
            ExcoCalendar calendar = new ExcoCalendar(DateTime.Today.Year - 2000, DateTime.Today.Month, false, plantID);
            // get customer, currency, territory, plant and year/period
            ExcoODBC database = ExcoODBC.Instance;
            database.Open(Database.CMSDAT);
            string query = "select dhinv#, dhbcs#, dhscs#, dhcurr, dhterr, dharyr, dharpr, dhord# from cmsdat.oih where dhpost='Y' and dharyr>=" + (calendar.GetFiscalYear()-2).ToString() + " and dhplnt=" + plantID.ToString();
            OdbcDataReader reader = database.RunQuery(query);
            while (reader.Read())
            {
                int invNum = Convert.ToInt32(reader["dhinv#"]);
                int soNum = Convert.ToInt32(reader["dhord#"]);
                string billTo = reader["dhbcs#"].ToString().Trim();
                string shipTo = reader["dhscs#"].ToString().Trim();
                string currency = reader["dhcurr"].ToString().Trim();
                string territory = reader["dhterr"].ToString().Trim();
                int invoiceYear = Convert.ToInt32(reader["dharyr"]);
                int invoicePeriod = Convert.ToInt32(reader["dharpr"]);
                ExcoCustomer customer;
                if (customerBillToMap.ContainsKey(billTo))
                {
                    customer = customerBillToMap[billTo];
                }
                else if (customerShipToMap.ContainsKey(billTo))
                {
                    customer = customerShipToMap[billTo];
                }
                else if (customerBillToMap.ContainsKey(shipTo))
                {
                    customer = customerBillToMap[shipTo];
                }
                else if (customerShipToMap.ContainsKey(shipTo))
                {
                    customer = customerShipToMap[shipTo];
                }
                else
                {
                    throw new Exception("Unknown customer: " + billTo + " " + shipTo + " invoice#:" + invNum.ToString());
                }
                Invoice invoice = new Invoice();
                invoice.orderNum = soNum;
                invoice.invoiceNum = invNum;
                invoice.customer = customer;
                invoice.currency = currency;
                invoice.territory = territory;
                invoice.plant = plantID;
                invoice.calendar = new ExcoCalendar(invoiceYear, invoicePeriod, true, plantID);
                invoiceMap.Add(invNum, invoice);
            }
            reader.Close();
            // get sales
            query = "select diinv#, sum(dipric*diqtsp) as sale from cmsdat.oid where diglcd='SAL' group by diinv#";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                int invNum = Convert.ToInt32(reader[0]);
                if (invoiceMap.ContainsKey(invNum))
                {
                    invoiceMap[invNum].sale += Convert.ToDouble(reader
[1]);
                }
            }
            reader.Close();
            // get discount
            query = "select flinv#, coalesce(sum(fldext), 0.0) from cmsdat.ois where fldisc like 'D%' or fldisc like 'M%' group by flinv#";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                int invNum = Convert.ToInt32(reader[0]);
                if (invoiceMap.ContainsKey(invNum))
                {
                    invoiceMap[invNum].discount += Convert.ToDouble(reader
[1]);
                }
            }
            reader.Close();
            // get fast track
            query = "select flinv#, coalesce(sum(fldext), 0.0) from cmsdat.ois where fldisc like 'F%' group by flinv#";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                int invNum = Convert.ToInt32(reader[0]);
                if (invoiceMap.ContainsKey(invNum))
                {
                    invoiceMap[invNum].fastTrack += Convert.ToDouble(reader[1]);
                }
            }
            reader.Close();
            // get steel surcharge
            query = "select flinv#, coalesce(sum(fldext), 0.0) from cmsdat.ois where fldisc like 'S%' or fldisc like 'P%' group by flinv#";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                int invNum = Convert.ToInt32(reader[0]);
                if (invoiceMap.ContainsKey(invNum))
                {
                    invoiceMap[invNum].surcharge += Convert.ToDouble(reader[1]);
                }
            }
            reader.Close();
            // get credit
            query = "select dhinv#, coalesce(dipric*(max(diqtso,diqtsp)), 0.0) from cmsdat.oih, cmsdat.oid where dhincr='C' and dhpost='Y' and dhinv#=diinv# and diglcd='SAL'";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                int invNum = Convert.ToInt32(reader[0]);
                if (invoiceMap.ContainsKey(invNum))
                {
                    invoiceMap[invNum].credit += Convert.ToDouble(reader[1]);
                }
            }
            reader.Close();
            // get budgets
            database.Open(Database.DECADE_MARKHAM);
            switch (plantID)
            {
                case 1:
                    query = "select CustomerID, Currency, Period01, Period02, Period03, Period04, Period05, Period06, Period07, Period08, Period09, Period10, Period11, Period12 from tiger.dbo.Markham_Budget where Year=20" + fiscalYear.ToString();
                    break;
                case 4:
                    query = "select CustomerID, Currency, Period01, Period02, Period03, Period04, Period05, Period06, Period07, Period08, Period09, Period10, Period11, Period12 from tiger.dbo.Colombia_Budget where Year=20" + fiscalYear.ToString();
                    break;
                case 3:
                    query = "select CustomerID, Currency, Period01, Period02, Period03, Period04, Period05, Period06, Period07, Period08, Period09, Period10, Period11, Period12 from tiger.dbo.Michigan_Budget where Year=20" + fiscalYear.ToString();
                    break;
                case 5:
                    query = "select CustomerID, Currency, Period01, Period02, Period03, Period04, Period05, Period06, Period07, Period08, Period09, Period10, Period11, Period12 from tiger.dbo.Texas_Budget where Year=20" + fiscalYear.ToString();
                    break;
            }
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                string currency = reader[1].ToString();
                List<ExcoMoney> budgetList = new List<ExcoMoney>();
                for (int i = 0; i < 12; i++)
                {
                    ExcoMoney budget = new ExcoMoney(new ExcoCalendar(fiscalYear, i + 1, true, 1), Convert.ToDouble(reader[i + 2]), currency);
                    budgetList.Add(budget);
                }
                budgetMap.Add(reader[0].ToString().Trim(), budgetList);
            }
            reader.Close();
            // get parts
            database.Open(Database.CMSDAT);
            query = "select diord#, dimajr, didesc, dipart from cmsdat.oid where distkl='" + plantID.ToString("D3") + "PRD'";
            reader = database.RunQuery(query);
            while (reader.Read())
            {
                int soNum = Convert.ToInt32(reader["diord#"]);
                string part = reader["dipart"].ToString().Trim();
                string desc = reader["didesc"].ToString().Trim();
                string type = reader["dimajr"].ToString().Trim();
                if (desc.Contains("NCR ") || part.Contains("NCR "))
                {
                    type = "NCR";
                }
                if (partMap.ContainsKey(soNum))
                {
                    partMap[soNum].Add(type);
                }
                else
                {
                    List<string> typeList = new List<string>();
                    typeList.Add(type);
                    partMap.Add(soNum, typeList);
                }
            }
            reader.Close();
            // build plant data structure
            foreach (ExcoCustomer customer in customerBillToMap.Values)
            {
                plant.AddCustomer(customer, invoiceMap, fiscalYear, doesIncludeSurcharge, budgetMap, partMap);
            }
            plant.GetSurcharge();
            plant.CountParts();
        }
    }
}
