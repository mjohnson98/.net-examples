using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using DAL;

namespace ExcelReader.Spreadsheets.NationalDefenseAndSupportForVeterans.DefenseSpending
{
    public class NationalDefenseSpending : SheetReader
    {
        public override InputSpreadsheet SeedSpreadsheet()
        {
            var sheet = new InputSpreadsheet()
            {
                ColumnNamesRow = 5,
                DataStartColumn = 1,
                Metric = "National defense spending",
                Path = @"C:\Source\Ballmer\jsonServer\Ballmer\Data",
                SheetName = "National Defense spend.xlsx",
                WorksheetIndex = 0,
                Extension = Enums.SpreadsheetExtension.xlsx,
                TopicName = "National Defense - Spending",
                Table = "National defense spending",
                Format = Enums.SpreadsheetFormat.TotalsAsHeader,
                HasMetricData = true,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };
           
            var range = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 6,
                ColumnEnd = 45,
                RowEnd = 6,
                IndentLevel = 0,
                Sequence = 1
            };
            sheet.DataRange = new List<InputSheetDataRange>();
            sheet.DataRange.Add(range); 
            
            #region type of spending
            var bySpendingType = new InputSpreadsheet()
            {
                Title = "By type",
                ChartableChildrenDescription = "By type",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };
            bySpendingType.DataRange = new List<InputSheetDataRange>();
            sheet.Breakdowns = new List<InputSpreadsheet>();
            #endregion

            #region expenditures
            var expendituresBreakdown = new InputSpreadsheet()
            {
                Title = "Expenditures",
                ChartableChildrenDescription = "Expenditures",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var expendituresBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 7,
                ColumnEnd = 45,
                RowEnd = 7,
                IndentLevel = 0,
                Sequence = 1
            };

            #region personnel compensation range
            var personnelCompensationRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 8,
                ColumnEnd = 45,
                RowEnd = 8,
                IndentLevel = 1,
                Sequence = 1
            };
            #region type of personnel compensation
            var typeOfPersonnelExpendituresBreakdown = new InputSpreadsheet()
            {
                Title = "Type of Personnel Compensation",
                ChartableChildrenDescription = "Type of Personnel Compensation",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var typeOfPersonnelExpendituresBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 9,
                ColumnEnd = 45,
                RowEnd = 10,
                IndentLevel = 1,
                Sequence = 1
            };
            typeOfPersonnelExpendituresBreakdown.DataRange = new List<InputSheetDataRange>();
            typeOfPersonnelExpendituresBreakdown.DataRange.Add(typeOfPersonnelExpendituresBreakdownRange);
            #endregion
            personnelCompensationRange.Breakdowns = new List<InputSpreadsheet>();
            personnelCompensationRange.Breakdowns.Add(typeOfPersonnelExpendituresBreakdown);
            #endregion
            expendituresBreakdown.DataRange = new List<InputSheetDataRange>();
            expendituresBreakdown.DataRange.Add(personnelCompensationRange);

            #region consumption of capital / depreciation range
            var capitalConsumptionRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 11,
                ColumnEnd = 45,
                RowEnd = 11,
                IndentLevel = 1,
                Sequence = 1
            };
            #endregion
            expendituresBreakdown.DataRange.Add(capitalConsumptionRange);

            #region durable goods purchased range
            var durableGoodsRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 12,
                ColumnEnd = 45,
                RowEnd = 12,
                IndentLevel = 1,
                Sequence = 1
            };
            #region type of durable goods purchased
            var typeOfDurableGoodsBreakdown = new InputSpreadsheet()
            {
                Title = "Type of Durable Goods Purchased",
                ChartableChildrenDescription = "Type of Durable Goods Purchased",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var typeOfDurableGoodsBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 13,
                ColumnEnd = 45,
                RowEnd = 18,
                IndentLevel = 1,
                Sequence = 1
            };
            typeOfDurableGoodsBreakdown.DataRange = new List<InputSheetDataRange>();
            typeOfDurableGoodsBreakdown.DataRange.Add(typeOfDurableGoodsBreakdownRange);
            #endregion
            durableGoodsRange.Breakdowns = new List<InputSpreadsheet>();
            durableGoodsRange.Breakdowns.Add(typeOfDurableGoodsBreakdown);
            #endregion
            expendituresBreakdown.DataRange.Add(durableGoodsRange);

            #region nondurable goods purchased range
            var nondurableGoodsRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 19,
                ColumnEnd = 45,
                RowEnd = 19,
                IndentLevel = 1,
                Sequence = 1
            };
            #region type of nondurable goods purchased
            var typeOfNondurableGoodsBreakdown = new InputSpreadsheet()
            {
                Title = "Type of Nondurable Goods Purchased",
                ChartableChildrenDescription = "Type of Nondurable Goods Purchased",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var typeOfNondurableGoodsBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 20,
                ColumnEnd = 45,
                RowEnd = 22,
                IndentLevel = 1,
                Sequence = 1
            };
            typeOfNondurableGoodsBreakdown.DataRange = new List<InputSheetDataRange>();
            typeOfNondurableGoodsBreakdown.DataRange.Add(typeOfNondurableGoodsBreakdownRange);
            #endregion
            nondurableGoodsRange.Breakdowns = new List<InputSpreadsheet>();
            nondurableGoodsRange.Breakdowns.Add(typeOfNondurableGoodsBreakdown);
            #endregion
            expendituresBreakdown.DataRange.Add(nondurableGoodsRange);

            #region services purchased range
            var servicesPurchasedRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 23,
                ColumnEnd = 45,
                RowEnd = 23,
                IndentLevel = 1,
                Sequence = 1
            };
            #region type of services purchased
            var typeOfServicesPurchasedBreakdown = new InputSpreadsheet()
            {
                Title = "Type of Services Purchased",
                ChartableChildrenDescription = "Type of Services Purchased",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var typeOfServicesPurchasedBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 24,
                ColumnEnd = 45,
                RowEnd = 28,
                IndentLevel = 1,
                Sequence = 1
            };
            typeOfServicesPurchasedBreakdown.DataRange = new List<InputSheetDataRange>();
            typeOfServicesPurchasedBreakdown.DataRange.Add(typeOfServicesPurchasedBreakdownRange);
            #endregion
            servicesPurchasedRange.Breakdowns = new List<InputSpreadsheet>();
            servicesPurchasedRange.Breakdowns.Add(typeOfServicesPurchasedBreakdown);
            #endregion
            expendituresBreakdown.DataRange.Add(servicesPurchasedRange);

            #region own-account investment / sales to other range
            var accountInvestmentRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 29,
                ColumnEnd = 45,
                RowEnd = 29,
                IndentLevel = 0,
                Sequence = 1
            };
            #endregion
            expendituresBreakdown.DataRange.Add(accountInvestmentRange);

            expendituresBreakdownRange.Breakdowns = new List<InputSpreadsheet>();
            expendituresBreakdownRange.Breakdowns.Add(expendituresBreakdown);

            bySpendingType.DataRange = new List<InputSheetDataRange>();
            bySpendingType.DataRange.Add(expendituresBreakdownRange);


            #endregion

            #region investment
            var investmentBreakdown = new InputSpreadsheet()
            {
                Title = "Investment",
                ChartableChildrenDescription = "Investment",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var investmentBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 32,
                ColumnEnd = 45,
                RowEnd = 32,
                IndentLevel = 0,
                Sequence = 1
            };

            #region structures range
            var structuresRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 33,
                ColumnEnd = 45,
                RowEnd = 33,
                IndentLevel = 0,
                Sequence = 1
            };

            #endregion
            investmentBreakdown.DataRange = new List<InputSheetDataRange>();
            investmentBreakdown.DataRange.Add(structuresRange);
            
            #region equipment range
            var equipmentRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 34,
                ColumnEnd = 45,
                RowEnd = 34,
                IndentLevel = 0,
                Sequence = 1
            };
            #region type of equipment
            var typeOfEquipmentBreakdown = new InputSpreadsheet()
            {
                Title = "Type of Equpiment",
                ChartableChildrenDescription = "Type of Equpiment",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var typeOfEquipmentBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 35,
                ColumnEnd = 45,
                RowEnd = 40,
                IndentLevel = 0,
                Sequence = 1
            };
            typeOfEquipmentBreakdown.DataRange = new List<InputSheetDataRange>();
            typeOfEquipmentBreakdown.DataRange.Add(typeOfEquipmentBreakdownRange);
            #endregion
            equipmentRange.Breakdowns = new List<InputSpreadsheet>();
            equipmentRange.Breakdowns.Add(typeOfEquipmentBreakdown);
            #endregion
            investmentBreakdown.DataRange.Add(equipmentRange);

            #region intellectual property range
            var intellectualPropertyRange = new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 41,
                ColumnEnd = 45,
                RowEnd = 41,
                IndentLevel = 0,
                Sequence = 1
            };
            #region type of intellectual property
            var typeOfIntellectualPropertyBreakdown = new InputSpreadsheet()
            {
                Title = "Type of Intellectual Property",
                ChartableChildrenDescription = "Type of Intellectual Property",
                Format = Enums.SpreadsheetFormat.NoTotal,
                HasMetricData = false,
                Units = 1000000000,
                RoundingUnit = 1000000000,
                SigFigs = 1,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.Dollars,
                MetricType = Enums.MetricType.Spending,
                IsPrimary = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true,
                Processed = false,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                SlideNumber = 152,
                SlideTitle = "National defense spending",
            };

            var typeOfIntellectualPropertyBreakdownRange= new InputSheetDataRange()
            {
                ColumnStart = 0,
                RowStart = 42,
                ColumnEnd = 45,
                RowEnd = 51,
                IndentLevel = 0,
                Sequence = 1
            };
            typeOfIntellectualPropertyBreakdown.DataRange = new List<InputSheetDataRange>();
            typeOfIntellectualPropertyBreakdown.DataRange.Add(typeOfIntellectualPropertyBreakdownRange);
            #endregion
            intellectualPropertyRange.Breakdowns = new List<InputSpreadsheet>();
            intellectualPropertyRange.Breakdowns.Add(typeOfIntellectualPropertyBreakdown);
            #endregion
            investmentBreakdown.DataRange.Add(intellectualPropertyRange);

            investmentBreakdownRange.Breakdowns = new List<InputSpreadsheet>();
            investmentBreakdownRange.Breakdowns.Add(investmentBreakdown);

            bySpendingType.DataRange.Add(investmentBreakdownRange);

            #endregion

            sheet.Breakdowns = new List<InputSpreadsheet>();
            sheet.Breakdowns.Add(bySpendingType);
            
            var footnoteRange = new InputSheetDataRange()
            {
                ColumnStart = 47,
                RowStart = 6,
                RowEnd = 9,
            };
            sheet.Footnote = footnoteRange;

            sheet.AdjustmentKeys = "i";
            return sheet;
        }

        public override void LoadSpreadsheet(InputSpreadsheet sheet)
        {
            base.LoadSpreadsheetHeader(sheet);
        }
    }
}
