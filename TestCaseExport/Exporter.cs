using System.Collections.Generic;
using System;
using System.Data;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.TeamFoundation.TestManagement.Client;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace TestCaseExport
{
    /// <summary>
    /// Exports a passed set of test cases to the supplied file.
    /// </summary>
    public class Exporter
    {

        int stepID;

        public void Export(string filename, ITestSuiteBase testSuite)
        {
            using (var pkg = new ExcelPackage())
            {
                
                int testCaseID = 1;
                
                
                foreach (var testCase in testSuite.AllTestCases)
                {

                    int row = 15;
                    stepID = 1;
                    int duplicateInstance = 0;
                    ExcelWorksheet sheet;

retry:
                    try
                    {
                        if(duplicateInstance == 0)
                        {
                            sheet = pkg.Workbook.Worksheets.Add(testCase.Id.ToString() + "." + testCase.Title);
                        }
                        else
                        {
                            sheet = pkg.Workbook.Worksheets.Add(testCase.Id.ToString() + "_" + duplicateInstance.ToString() + "." + testCase.Title);
                        }
                    }
                    catch
                    {
                        duplicateInstance++;
                        goto retry;
                    }
                    

                                        

                    #region Sheet Formatting

                    //General Sheet Formatting
                    sheet.Column(1).Width = 1.86;
                    sheet.Column(2).Width = 7.29;
                    sheet.Column(3).Width = 35;
                    sheet.Column(4).Width = 37;
                    sheet.Column(5).Width = 37;
                    sheet.Column(6).Width = 1.43;

                    sheet.Row(7).Height = 21.75;
                    sheet.Row(53).Height = 9;

                    sheet.Cells[15, 2, 34, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[15, 2, 34, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    //Range B8 - E9
                    sheet.Cells[8, 2, 8, 5].Merge = true;
                    sheet.Cells[9, 2, 9, 5].Merge = true;
                    sheet.Cells[8, 2].Style.Font.Bold = true;
                    sheet.Cells[8, 2].Value = "Test Case ID";
                    sheet.Cells[9, 2].Value = testCase.WorkItem.Id + ". " + testCase.WorkItem.Title;

                    //Range B11 - B12
                    sheet.Cells[11, 2, 11, 5].Merge = true;
                    sheet.Cells[12, 2, 12, 5].Merge = true;
                    sheet.Cells[11, 2].Style.Font.Bold = true;
                    sheet.Cells[11, 2].Value = "Test Item:";
                    sheet.Cells[12, 2].Value = testCase.WorkItem.Description;

                    //Range E2 - E5
                    sheet.Cells[2, 5].Value = "Doc No: RDI_CRB_INT_TXT2016_0027r01";
                    sheet.Cells[3, 5].Value = "Rev No: 01";
                    sheet.Cells[4, 5].Value = "Rev Date: 10-06-2016";
                    sheet.Cells[5, 5].Value = "Page: 1 of 1";
                    sheet.Cells[2, 5, 5, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Justify;

                    //Range D3 - D4
                    sheet.Cells[3, 4, 4, 4].Merge = true;
                    sheet.Cells[3, 4].Value = testCase.Area + " TEST CASE";
                    sheet.Cells[3, 4].Style.Font.Size = 14;
                    sheet.Cells[3, 4].Style.Font.Bold = true;
                    sheet.Cells[3, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[3, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                    //Range B2 - B7
                    sheet.Cells[7, 2].Value = testCase.WorkItem.Title;
                    sheet.Cells[7, 2, 7, 5].Merge = true;
                    sheet.Cells[7, 2, 7, 5].Style.Font.Bold = true;
                    sheet.Cells[7, 2].Style.Font.Size = 18;
                    sheet.Cells[7, 2, 7, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[7, 2, 7, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    
                    // B14 - E14
                    sheet.Cells[14, 2].Value = "No";
                    sheet.Cells[14, 3].Value = "Input Specifications";
                    sheet.Cells[14, 4].Value = "Output Specifications";
                    sheet.Cells[14, 5].Value = "Exceptions";
                    sheet.Cells[14, 2, 14, 5].Style.Font.Bold = true;
                    sheet.Cells[14, 2, 14, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[14, 2, 14, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    //B36 - E37
                    sheet.Cells[36, 2, 36, 5].Merge = true;
                    sheet.Cells[37, 2, 37, 5].Merge = true;
                    sheet.Cells[38, 2, 38, 5].Merge = true;
                    sheet.Cells[39, 2, 39, 5].Merge = true;
                    sheet.Cells[40, 2, 40, 5].Merge = true;
                    sheet.Cells[36, 2, 36, 5].Style.Font.Bold = true;
                    sheet.Cells[36, 2].Value = "Enviromental Needs";
                    sheet.Cells[37, 2].Value = "Software and Hardware Requirements";

                    //B42 - E46
                    sheet.Cells[42, 2, 42, 5].Merge = true;
                    sheet.Cells[43, 2, 43, 5].Merge = true;
                    sheet.Cells[44, 2, 44, 5].Merge = true;
                    sheet.Cells[45, 2, 45, 5].Merge = true;
                    sheet.Cells[46, 2, 46, 5].Merge = true;
                    sheet.Cells[42, 2, 42, 5].Style.Font.Bold = true;
                    sheet.Cells[42, 2].Value = "Special Procedural Requirements";
                    sheet.Cells[43, 2].Value = "Constraints etc";

                    //B48 - E52
                    sheet.Cells[48, 2, 48, 5].Merge = true;
                    sheet.Cells[49, 2, 49, 5].Merge = true; 
                    sheet.Cells[50, 2, 50, 5].Merge = true; 
                    sheet.Cells[51, 2, 51, 5].Merge = true; 
                    sheet.Cells[52, 2, 52, 5].Merge = true;
                    sheet.Cells[48, 2, 48, 5].Style.Font.Bold = true;
                    sheet.Cells[48, 2].Value = "Intercase Dependencies";
                    sheet.Cells[49, 2].Value = "List of prerequisites or inter-acting test cases";

                    //Border Formatting
                    sheet.Cells[7, 2, 7, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[2, 2, 5, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[2, 4, 5, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[8, 2, 9, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[11, 2, 12, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[14, 2, 34, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[14, 3, 34, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[14, 4, 34, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[14, 5, 34, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[14, 2, 14, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[36, 2, 36, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[37, 2, 40, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[42, 2, 42, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[42, 2, 46, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[48, 2, 48, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[49, 2, 52, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                 
                    sheet.Cells[1, 1, 53, 6].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    //var header = sheet.Cells[1, 1, 1, 8];
                    //header.Style.Font.Bold = true;
                    //header.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    //header.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 226, 238, 18));

                    #endregion

                    testCaseID++;

                        var replacementSets = GetReplacementSets(testCase);
                            foreach (var replacements in replacementSets)
                            {
                                var firstRow = row;
                                foreach (var testAction in testCase.Actions)
                                {
                                    CellFormatting(ref stepID, sheet, ref row);
                                    AddSteps(sheet, testAction, replacements, ref row); 
                                }
                                //if (firstRow != row)
                                //{
                                //    //var mergedID = sheet.Cells[firstRow, 1, row - 1, 1];
                                //    //mergedID.Merge = true;
                                //    ////mergedID.Value = testCase.WorkItem == null ? "" : testCase.WorkItem.Id.ToString();
                                //    //mergedID.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //    //mergedID.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                //    //mergedID.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                //    //var mergedText = sheet.Cells[firstRow, 2, row - 1, 2];
                                //    //mergedText.Merge = true;
                                //    ////CleanupText(mergedText, testCase.Title, replacements);
                                //    //mergedText.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //    //mergedText.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                                //}
                            }
                        }

                        pkg.SaveAs(new FileInfo(filename));
            }
        }

        private List<Dictionary<string, string>> GetReplacementSets(ITestCase testCase)
        {
            var replacementSets = new List<Dictionary<string, string>>();
            foreach (DataRow r in testCase.DefaultTableReadOnly.Rows)
            {
                var replacement = new Dictionary<string, string>();
                foreach (DataColumn c in testCase.DefaultTableReadOnly.Columns)
                {
                    replacement[c.ColumnName] = r[c] as string;
                }
                replacementSets.Add(replacement);
            }
            return replacementSets.DefaultIfEmpty(new Dictionary<string, string>()).ToList();
        }

        private void AddSteps(ExcelWorksheet xlWorkSheet, ITestAction testAction, Dictionary<string, string> replacements, ref int row)
        {
            var testStep = testAction as ITestStep;
            var group = testAction as ITestActionGroup;
            var sharedRef = testAction as ISharedStepReference;
            if (null != testStep)
            {
                CleanupText(xlWorkSheet.Cells[row, 3], testStep.Title.ToString(), replacements);
                CleanupText(xlWorkSheet.Cells[row, 4], testStep.ExpectedResult.ToString(), replacements);
            }
            else if (null != group)
            {
                //foreach (var action in group.Actions)
                //{
                //    AddSteps(xlWorkSheet, action, replacements, ref row);
                //}
            }
            else if (null != sharedRef)
            {
                var step = sharedRef.FindSharedStep();
                foreach (var action in step.Actions)
                {
                    AddSteps(xlWorkSheet, action, replacements, ref row);
                }
            }
            row++;
        }

        private void CleanupText(ExcelRangeBase cell, string input, Dictionary<string, string> replacements)
        {
            foreach (var kvp in replacements)
            {
                input = input.Replace("@" + kvp.Key, kvp.Value);
            }

            new HtmlToRichTextHelper().HtmlToRichText(cell, input);
        }

        private void CellFormatting(ref int stepID,  ExcelWorksheet sheet, ref int row)
        {
            sheet.Cells[row, 3, row, 4].Style.WrapText = true;
            sheet.Cells[row, 2].Value = stepID; 
            stepID++;
        }
       
    }
}
