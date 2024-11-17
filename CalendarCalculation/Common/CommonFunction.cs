using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using EXCEL = Microsoft.Office.Interop.Excel;

namespace CalendarCalculation.Common
{
    public static class CommonFunction
    {
        public static void ShowNoticeDialog(string message)
        {
            MessageBox.Show(message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ShowWarningDialog(string message)
        {
            MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void ShowErrorDialog(string message)
        {
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static DialogResult ShowQuestionDialog(string message)
        {
            return MessageBox.Show(message, "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        public static bool IsHiddenFile(string filePath)
        {
            FileAttributes attributes = File.GetAttributes(filePath);
            return (attributes & FileAttributes.Hidden) == FileAttributes.Hidden;
        }

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        public static void KillExcel(EXCEL.Application theApp)
        {
            int id = 0;
            IntPtr intptr = new IntPtr(theApp.Hwnd);
            System.Diagnostics.Process p = null;
            try
            {
                GetWindowThreadProcessId(intptr, out id);
                p = System.Diagnostics.Process.GetProcessById(id);
                if (p != null)
                {
                    p.Kill();
                    p.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("KillExcel:" + ex.Message);
            }
        }

        public static int GetLastRow(ExcelWorksheet ws, int colNumber)
        {
            return ws.Cells.Last(c => c.Start.Column == colNumber).End.Row;
        }

        public static string Setvalue(int lineNumber, string abnormalSheetName, string defaultAbnormalSheetName)
        {
            string currentPath = Directory.GetCurrentDirectory();
            string filePath = Path.Combine(currentPath, "sheetName.txt");

            if (string.IsNullOrEmpty(ReadLineFromFile(filePath, lineNumber))) return defaultAbnormalSheetName;
            else
            {
                abnormalSheetName = ReadLineFromFile(filePath, lineNumber);

            } 
            return abnormalSheetName;
        }

        public static string ReadLineFromFile(string fileName, int lineNumber)
        {
            if (lineNumber < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(lineNumber), "Line number must be greater than 0.");
            }

            try
            {
                using (var reader = new StreamReader(fileName))
                {
                    int currentLine = 0;
                    string line;

                    while ((line = reader.ReadLine()) != null)
                    {
                        currentLine++;
                        if (currentLine == lineNumber)
                        {
                            return line;
                        }
                    }

                    throw new ArgumentOutOfRangeException(nameof(lineNumber), "Line number exceeds the total number of lines in the file.");
                }
            }
            catch (FileNotFoundException)
            {
                throw new FileNotFoundException($"The file '{fileName}' was not found.");
            }
            catch (IOException ex)
            {
                throw new IOException("An error occurred while reading the file.", ex);
            }
        }
    }
}
