using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

//
// This application depends on iText NuGet packages. Please install the following packages:
//
//      Install-Package itext
//      Install-Package itext7.bouncy-castle-adapter
//
namespace AMC10EmailSender
{
    internal class Program
    {
        // This is the file downloaded from MAA Portal, with all students' score in it
        private const string ScorePdfPath = @"C:\Documents\TestDetailsByName.aspx.pdf";

        // This is the file downloaded from Math Club portal, with all students' contact info in it
        private const string StudentContactsFile = @"C:\Documents\Competitions & Teams.csv";

        static void Main(string[] args)
        {
            // Crete a folder to contain each student's score file and email message file
            string studentScoresDir = string.Format(CultureInfo.InvariantCulture, "{0}\\StudentScores", Path.GetDirectoryName(ScorePdfPath));

            if (!Directory.Exists(studentScoresDir))
            {
                Directory.CreateDirectory(studentScoresDir);
            }

            // Create individual score files for each student so that they don't see each other's scores
            ScoreFileHandler pdfHandler = new ScoreFileHandler();

            pdfHandler.SplitScorePdfFile(ScorePdfPath, studentScoresDir);

            // Create email template for each student
            List<StudentContact> studentContacts = ReadStudentContacts();

            List<Tuple<string, string>> studentScores = pdfHandler.GetStudentScores(ScorePdfPath);

            foreach(Tuple<string, string> studentScore in studentScores)
            {
                // studentScore.Item1 is in the format of "LastName, FirstName"
                string[] studentNameParts = studentScore.Item1.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                string firstName = studentNameParts[1].Trim();
                string lastName = studentNameParts[0].Trim();

                string studentName = string.Format(CultureInfo.InvariantCulture, "{0} {1}", firstName, lastName);

                StudentContact studentContact = 
                    studentContacts.Find(s => s.FirstName.Equals(firstName, StringComparison.OrdinalIgnoreCase) && s.LastName.Equals(lastName, StringComparison.OrdinalIgnoreCase));

                if (studentContact == null)
                {
                    Console.WriteLine("Error: Could not find contact infomation for student {0}", studentName);

                    continue;
                }

                string emailMessage = CreateEmailMessage(studentContact, studentScore.Item2);

                string emailFilePath = string.Format(CultureInfo.InvariantCulture, @"{0}\{1}.txt", studentScoresDir, studentScore.Item1);
                File.WriteAllText(emailFilePath, emailMessage);
            }
        }

        private static List<StudentContact> ReadStudentContacts()
        {
            List<StudentContact> studentContacts = new List<StudentContact>();

            using (StreamReader sr = new StreamReader(StudentContactsFile))
            {
                // Skip header
                string line = sr.ReadLine();

                while ((line = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    // In this student contact file, the columns are (in this order):
                    //      Student Name, Grade, Student Email, Parent Email
                    string[] lineParts = line.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (!lineParts.Any())
                    {
                        continue;
                    }

                    string studentName = lineParts[0].Trim();
                    string[] studentNameParts = studentName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    string studentEmail = lineParts[2].Trim();
                    string parentEmail = lineParts[3].Trim();

                    studentContacts.Add(new StudentContact()
                    {
                        FirstName = studentNameParts.First(),
                        LastName = studentNameParts.Last(),
                        StudentEmail = studentEmail,
                        ParentEmail = parentEmail
                    });
                }
            }

            return studentContacts;
        }

        private static string CreateEmailMessage(StudentContact contact, string studentScore)
        {
            StringBuilder message = new StringBuilder();

            message.AppendLine(string.Format(CultureInfo.InvariantCulture, "{0},{1}", contact.StudentEmail, contact.ParentEmail));
            message.AppendLine();

            message.AppendLine(string.Format(CultureInfo.InvariantCulture, "{0}, {1} - Fall 2023 AMC 10A Result", contact.LastName, contact.FirstName));
            message.AppendLine();

            string messageBodyTemplate = @"Hi {0},

Your Fall 2023 AMC 10A Score is {1}. Please find your score details in the attached file.

In the attached file, you can also see the correction rate of each problem. This is the statistics of all 15 participants from Redmond Middle School.

We haven't received the AIME qualifier list. Once we receive it, we will notifier each qualifier separately.

Please let us know if you have any questions.

RMS Math Club";

            string messageBody = string.Format(CultureInfo.InvariantCulture, messageBodyTemplate, contact.FirstName, studentScore);

            message.AppendLine(messageBody);

            return message.ToString();
        }
    }
}
