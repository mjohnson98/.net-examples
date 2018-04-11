using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using DAL;

namespace ExcelReader.Spreadsheets.EnsureJustice.CrimeAndJustice.CrimeAndPolice
{
    public class Arrests : SheetReader
    {
        public override InputSpreadsheet SeedSpreadsheet()
        {
            var sheet = new InputSpreadsheet()
            {
                ColumnNamesRow = 2,
                DataStartColumn = 2,
                Metric = "Arrests",
                Path = @"C:\Source\Ballmer\jsonServer\Ballmer\Data",
                SheetName = "Arrests.xlsx",
                WorksheetIndex = 0,
                Extension = Enums.SpreadsheetExtension.xlsx,
                TopicName = "Crime and police",
                Table = "Arrests",
                ChartableChildrenDescription = "Total arrests",
                Format = Enums.SpreadsheetFormat.TotalsAsHeader,
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                IsPrimary = true,
                HasMetricData = true,
                SheetReaderClass = this.GetType().Name,
                IsFeatured = true, 
                Processed = false,
                IsVisible = true,
                ApplicationType = Enums.USAFactsApplicationType.Missions,
                InternalId = new IdentityLink(),
                Definition = "The FBI’s Uniform Crime Reporting(UCR) Program counts one arrest for each separate instance in which a person is arrested, cited, or summoned for an offense. The UCR Program collects arrest data on 28 offenses, as described in Offense Definitions. (Please note that, beginning in 2010, the UCR Program no longer collected data on runaways.) Because a person may be arrested multiple times during a year, the UCR arrest figures do not reflect the number of individuals who have been arrested; rather, the arrest data show the number of times that persons are arrested, as reported by law enforcement agencies to the UCR Program.",
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var range = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 3,
                ColumnEnd = 39,
                RowEnd = 3,
                IndentLevel = 1
            };

            #region offense
            var offenseBreakdown = new InputSpreadsheet()
            {
                Title = "By offense",
                Format = Enums.SpreadsheetFormat.NoTotal,
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                IsPrimary = false,
                IsFeatured = false,
                HasMetricData = false,
                IsVisible = true,
                InternalId = new IdentityLink(),
                SheetReaderClass = this.GetType().Name,
                TopicBreakdowns = Enums.TopicBreakdowns.Type,
                Sequence = 1,
                ChartableChildrenDescription = "By offense",
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var offenseRange = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 4,
                ColumnEnd = 39,
                RowEnd = 36,
                IndentLevel = 1,
                IgnoredRows = "5,6,10,11,21"
            };
            offenseBreakdown.DataRange = new List<InputSheetDataRange>();
            offenseBreakdown.DataRange.Add(offenseRange);
            #region Arrests by Drug Abuse Violation
            var drugAbuseViolation = new InputSpreadsheet()
            {
                Title = "By drug abuse violation",
                Format = Enums.SpreadsheetFormat.NoTotal,
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                IsPrimary = false,
                IsFeatured = false,
                HasMetricData = false,
                IsVisible = true,
                InternalId = new IdentityLink(),
                SheetReaderClass = this.GetType().Name,
                TopicBreakdowns = Enums.TopicBreakdowns.Type,
                Sequence = 1,
                ChartableChildrenDescription = "By drug abuse violation",
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var drugAbuseViolationRange = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 5,
                ColumnEnd = 39,
                RowEnd = 6,
                IndentLevel = 2,
            };
            drugAbuseViolation.DataRange = new List<InputSheetDataRange>();
            drugAbuseViolation.DataRange.Add(drugAbuseViolationRange);
            #endregion
            offenseBreakdown.Breakdowns = new List<InputSpreadsheet>();
            offenseBreakdown.Breakdowns.Add(drugAbuseViolation);

            #region Arrest By Assault Type
            var assaultType = new InputSpreadsheet()
            {
                Title = "By assault type",
                Format = Enums.SpreadsheetFormat.NoTotal,
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                IsPrimary = false,
                IsFeatured = false,
                HasMetricData = false,
                IsVisible = true,
                InternalId = new IdentityLink(),
                SheetReaderClass = this.GetType().Name,
                TopicBreakdowns = Enums.TopicBreakdowns.Type,
                Sequence = 1,
                ChartableChildrenDescription = "By assault type",
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var assaultTypeRange = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 10,
                ColumnEnd = 39,
                RowEnd = 11,
                IndentLevel = 2,
            };
            assaultType.DataRange = new List<InputSheetDataRange>();
            assaultType.DataRange.Add(assaultTypeRange);
            #endregion
            offenseBreakdown.Breakdowns.Add(assaultType);
            #endregion
            range.Breakdowns = new List<InputSpreadsheet>();
            range.Breakdowns.Add(offenseBreakdown);

            #region race
            var raceBreakdown = new InputSpreadsheet()
            {
                Title = "By race",
                Format = Enums.SpreadsheetFormat.NoTotal,
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                IsPrimary = false,
                HasMetricData = false,
                IsFeatured = false,
                IsVisible = true,
                InternalId = new IdentityLink(),
                SheetReaderClass = this.GetType().Name,
                TopicBreakdowns = Enums.TopicBreakdowns.Race,
                Sequence = 1,
                ChartableChildrenDescription = "By race",
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var raceRange = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 41,
                ColumnEnd = 39,
                RowEnd = 44,
                IndentLevel = 1
            };
            raceBreakdown.DataRange = new List<InputSheetDataRange>();
            raceBreakdown.DataRange.Add(raceRange);
            #endregion
            range.Breakdowns.Add(raceBreakdown);

            #region under 18 breakdown
            var ageBreakdown = new InputSpreadsheet()
            {
                Title = "By age",
                Format = Enums.SpreadsheetFormat.NoTotal,
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                InternalId = new IdentityLink(),
                IsPrimary = false,
                HasMetricData = false,
                IsFeatured = false,
                IsVisible = true,
                SheetReaderClass = this.GetType().Name,
                TopicBreakdowns = Enums.TopicBreakdowns.Age,
                Sequence = 2,
                ChartableChildrenDescription = "By age",
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var under18Range = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 45,
                ColumnEnd = 39,
                RowEnd = 45,
                Sequence = 1,
                IndentLevel = 0
            };
            var over18Range = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 50,
                ColumnEnd = 36,
                RowEnd = 50,
                Sequence = 3,
                IndentLevel = 1
            };
            ageBreakdown.DataRange = new List<InputSheetDataRange>();
            ageBreakdown.DataRange.Add(under18Range);
            ageBreakdown.DataRange.Add(over18Range);

            #region under 18 by race
            var under18RaceBreakdown = new InputSpreadsheet()
            {
                InternalId = new IdentityLink(),
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                IsPrimary = false,
                HasMetricData = false,
                IsFeatured = false,
                IsVisible = true,
                TopicBreakdowns = Enums.TopicBreakdowns.Race,
                Title = "Under 18 By Race",
                ChartableChildrenDescription = "Under 18 By Race",
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var under18RaceBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 46,
                ColumnEnd = 36,
                RowEnd = 49,
                Sequence = 1,
                IndentLevel = 2
            };
            under18RaceBreakdown.DataRange = new List<InputSheetDataRange>();
            under18RaceBreakdown.DataRange.Add(under18RaceBreakdownRange);
            #endregion
            ageBreakdown.Breakdowns = new List<InputSpreadsheet>();
            ageBreakdown.Breakdowns.Add(under18RaceBreakdown);

            #region over 18 breakdown
            var over18RaceBreakdown = new InputSpreadsheet()
            {
                InternalId = new IdentityLink(),
                Units = 1,
                RoundingUnit = 1,
                SigFigs = 0,
                XAxis = Enums.USAFactsDataType.Years,
                YAxis = Enums.USAFactsDataType.People,
                MetricType = Enums.MetricType.Other,
                TopicBreakdowns = Enums.TopicBreakdowns.Race,
                Title = "Over 18 By Race",
                ChartableChildrenDescription = "Over 18 By Race",
                IsPrimary = false,
                HasMetricData = false,
                IsFeatured = false,
                IsVisible = true,
                ChartType = Enums.ChartType.Line,
                CanToggleStackedArea = false,
                DefaultStackedArea = false,
                StartsAtZero = false,
                SlideNumber = ArrestsCommon.SlideNumber,
                SlideTitle = ArrestsCommon.SlideTitle
            };
            var over18RaceBreakdownRange = new InputSheetDataRange()
            {
                ColumnStart = 1,
                RowStart = 51,
                ColumnEnd = 36,
                RowEnd = 54,
                Sequence = 1,
                IndentLevel = 2
            };
            over18RaceBreakdown.DataRange = new List<InputSheetDataRange>();
            over18RaceBreakdown.DataRange.Add(over18RaceBreakdownRange);
            #endregion
            ageBreakdown.Breakdowns.Add(over18RaceBreakdown);
            #endregion
            range.Breakdowns.Add(ageBreakdown);

            sheet.DataRange = new List<InputSheetDataRange>();
            sheet.DataRange.Add(range);
            return sheet;
        }

        public override void LoadSpreadsheet(InputSpreadsheet sheet) 
        {
            base.LoadSpreadsheetHeader(sheet);
        }
    }
}
