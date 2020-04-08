using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace PxlHandler
{
    internal class PxlProtocolHelper
    {
        public static bool JumpToCell(PxlPath pxlPath)
        {
            var fileName = pxlPath.Filename;
            var worksheetName = pxlPath.WorksheetName;
            var rangeAddress = pxlPath.RangeAddress;
            Application excelApp;
            try
            {
                excelApp = Marshal.GetActiveObject("Excel.Application") as Application;
                if (excelApp == null)
                {
                    Console.WriteLine("This should not happen. Marshal.GetActiveObject(\"Excel.Application\") is null.");
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                excelApp = new Application();
            }

            excelApp.Visible = true;
            excelApp.UserControl = false;
            try
            {
                if (!TryActivateWorkbookByName(excelApp, Path.GetFileName(fileName)))
                {
                    if (Path.IsPathRooted(fileName))
                    {
                        try
                        {
                            excelApp.Workbooks.Open(fileName);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                        }
                    }
                }

                Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    Console.WriteLine($"Expected to find {fileName}, but there is no active workbook.");
                    ShowWarningMessage("workbook", fileName);
                    return false;
                }

                if (workbook.Name.StripKnownExcelExtension() != Path.GetFileName(fileName).StripKnownExcelExtension())
                {
                    Console.WriteLine($"Expected to find {fileName}, but found workbook {workbook.Name}.");
                    ShowWarningMessage("workbook", fileName);
                    return false;
                }

                Sheets sheets = workbook.Sheets;
                Worksheet worksheet;
                try
                {
                    worksheet = sheets[worksheetName] as Worksheet;
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Could not get worksheet \"{worksheetName}\". {e.Message}");
                    ShowWarningMessage("worksheet", worksheetName);
                    return false;
                }

                if (worksheet == null)
                {
                    return false;
                }

                if (worksheet.Visible != XlSheetVisibility.xlSheetVisible)
                {
                    if (UserWantsSheetToStayHidden(worksheetName))
                    {
                        return false;
                    }

                    worksheet.Visible = XlSheetVisibility.xlSheetVisible;
                }

                worksheet.Activate();
                if (string.IsNullOrEmpty(rangeAddress))
                {
                    return true;
                }

                try
                {
                    excelApp.Goto(worksheet.Range[rangeAddress]);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Could not select range \"{worksheetName}\". {e.Message}");
                    ShowWarningMessage("range", rangeAddress);
                    return false;
                }

                return true;
            }
            finally
            {
                try
                {
                    excelApp.Visible = true;
                    excelApp.UserControl = true;
                }
                catch
                {
                    // pass
                }
            }
        }

        private static void ShowWarningMessage(string objectType, string objectName)
        {
            var text = $"Target {objectType} \"{objectName}\" could not be not found.";
            Console.WriteLine(text);
        }

        private static bool TryActivateWorkbookByName(Application excelApp, string workbookName)
        {
            var workbooksList = excelApp.Workbooks.OfType<Workbook>().Select(x => x.Name).ToArray();
            foreach (var s in new[] {workbookName, workbookName.StripKnownExcelExtension()})
            {
                try
                {
                    if (!workbooksList.Contains(s))
                    {
                        continue;
                    }

                    excelApp.Workbooks[s].Activate();
                    return true;
                }
                catch
                {
                    // pass
                }
            }

            return false;
        }

        private static bool UserWantsSheetToStayHidden(string worksheetName)
        {
            var text = $"Target worksheet \"{worksheetName}\" is hidden. Unhide?";
            Console.WriteLine(text);
            var keyChar = Console.ReadKey().KeyChar;
            return keyChar == 'Y' || keyChar == 'y';
        }
    }
}