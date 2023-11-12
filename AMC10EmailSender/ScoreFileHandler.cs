using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AMC10EmailSender
{
    internal class ScoreFileHandler
    {
        public static Tuple<string, string> GetStudentScoreFromPage(PdfDocument scorePdf, int pageNum)
        {
            string studentName = null;
            string studentScore = null;

            ITextExtractionStrategy its = new SimpleTextExtractionStrategy();
            string s = PdfTextExtractor.GetTextFromPage(scorePdf.GetPage(pageNum), its);

            string[] lines = s.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

            for (int index = 0; index < lines.Length; index++)
            {
                string line = lines[index].Trim();

                if (line.Contains("Student ID") && line.Contains("Student Name"))
                {
                    // Example: 12795 Badrish, Advait 8 11/8/2023
                    string studentNameLine = lines[index + 1].Trim();

                    string pattern = @"^\d+\s(?<StudentName>\w+,\s\w+)\s";
                    MatchCollection collection = Regex.Matches(studentNameLine, pattern);

                    studentName = collection[0].Groups["StudentName"].Value;

                    index++;
                }
                else if (line.Contains("Score"))
                {
                    studentScore = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[1].Trim();
                }
            }

            return new Tuple<string, string>(studentName, studentScore);
        }

        public List<Tuple<string, string>> GetStudentScores(string scorePdfFilePath)
        {
            List<Tuple<string, string>> studentScores = new List<Tuple<string, string>>();

            PdfDocument pdfDoc = new PdfDocument(new PdfReader(scorePdfFilePath));

            for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
            {
                Tuple<string, string> studentScore = GetStudentScoreFromPage(pdfDoc, pageNum);

                studentScores.Add(studentScore);
            }

            return studentScores;
        }

        public void SplitScorePdfFile(string scorePdfFilePath, string studentScoresDir)
        {
            string splitFilePath = string.Format(CultureInfo.InvariantCulture, "{0}\\{{0}}.pdf", studentScoresDir);

            PdfDocument pdfDoc = new PdfDocument(new PdfReader(scorePdfFilePath));

            List<int> pageNums = new List<int>();
            for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
            {
                pageNums.Add(pageNum);
            }

            IList<PdfDocument> splitDocs = new CustomPdfSplitter(pdfDoc, splitFilePath).SplitByPageNumbers(pageNums);

            foreach (PdfDocument splitDoc in splitDocs)
            {
                splitDoc.Close();
            }

            pdfDoc.Close();
        }

        private class CustomPdfSplitter : PdfSplitter
        {
            private PdfDocument parentPdf;
            private string splitFilePathFormat;
            private int pageNumber = 1;

            public CustomPdfSplitter(PdfDocument pdfDocument, string splitFilePathFormat) : base(pdfDocument)
            {
                this.parentPdf = pdfDocument;

                this.splitFilePathFormat = splitFilePathFormat;
            }

            protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
            {
                string studentName = ScoreFileHandler.GetStudentScoreFromPage(this.parentPdf, this.pageNumber).Item1;

                this.pageNumber++;

                return new PdfWriter(string.Format(CultureInfo.InvariantCulture, this.splitFilePathFormat, studentName));
            }
        }
    }
}
