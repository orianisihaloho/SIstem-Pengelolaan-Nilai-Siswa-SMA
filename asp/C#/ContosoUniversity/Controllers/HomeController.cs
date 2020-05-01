using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ContosoUniversity.DAL;
using ContosoUniversity.ViewModels;
using ContosoUniversity.Models;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ContosoUniversity.ControllersZ
{
    public class HomeController : Controller
    {
        private SchoolContext db = new SchoolContext();
        DataModel __context = new DataModel();
        public ActionResult Index()
        {
            var students = from s in db.Students
                           select s;

            var senrollments = from e in db.Enrollments
                           select e;
            return View(students);
        }

        public ActionResult About()
        {
            // Commenting out LINQ to show how to do the same thing in SQL.
            //IQueryable<EnrollmentDateGroup> = from student in db.Students
            //           group student by student.EnrollmentDate into dateGroup
            //           select new EnrollmentDateGroup()
            //           {
            //               EnrollmentDate = dateGroup.Key,
            //               StudentCount = dateGroup.Count()
            //           };

            // SQL version of the above LINQ code.
            string query = "SELECT EnrollmentDate, COUNT(*) AS StudentCount "
                + "FROM Person "
                + "WHERE Discriminator = 'Student' "
                + "GROUP BY EnrollmentDate";
            IEnumerable<EnrollmentDateGroup> data = db.Database.SqlQuery<EnrollmentDateGroup>(query);

            return View(data.ToList());
        }
        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        // Make PDF Student Grade
        public ActionResult FileResultCreatePdfStudentGrade()
        {
            MemoryStream workStream = new MemoryStream();
            StringBuilder status = new StringBuilder("");
            DateTime dTime = DateTime.Now;
            //file name to be created   
            string strPDFFileName = string.Format("SamplePdf" + dTime.ToString("yyyyMMdd") + "-" + ".pdf");
            Document doc = new Document();
            doc.SetMargins(10f, 10f, 10f, 10f);
            //Create PDF Table with 5 columns  
            PdfPTable tableLayout = new PdfPTable(2);
            doc.SetMargins(10f, 10f, 10f, 10f);
            //Create PDF Table  

            //file will created in this path  
            string strAttachment = Server.MapPath("~/Downloadss/" + strPDFFileName);


            PdfWriter.GetInstance(doc, workStream).CloseStream = false;
            doc.Open();

            //Add Content to PDF   
            doc.Add(Add_Content_To_PDF(tableLayout));

            // Closing the document  
            doc.Close();

            byte[] byteInfo = workStream.ToArray();
            workStream.Write(byteInfo, 0, byteInfo.Length);
            workStream.Position = 0;


            return File(workStream, "application/pdf", strPDFFileName);

        }

        protected PdfPTable Add_Content_To_PDF(PdfPTable tableLayout)
        {
           

            float[] headers = { 24, 45}; //Header Widths  
            tableLayout.SetWidths(headers); //Set the pdf headers  
            tableLayout.WidthPercentage = 100; //Set the PDF File witdh percentage  
            tableLayout.HeaderRows = 1;
            //Add Title to the PDF file at the top  

            var students = from s in db.Students
                           select s;

            var enrollments = from e in db.Enrollments
                               select e;

            tableLayout.AddCell(new PdfPCell(new Phrase("Laporan Nilai Siswa", new Font(Font.FontFamily.HELVETICA, 20, 1, new iTextSharp.text.BaseColor(0, 0, 0))))
            {
                Colspan = 12,
                Border = 0,
                PaddingBottom = 4,
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            foreach (var emp in students)
            {

                ViewBag.NameSortParm = "ID";
                AddCellToBody(tableLayout, emp.ID.ToString());
                ViewBag.NameSortParm = "Fullname";
                AddCellToBody(tableLayout, emp.FullName);

                ////Add header  
                /*  AddCellToHeader(tableLayout, "FullName");
                  AddCellToHeader(tableLayout, "Grade");
                  AddCellToHeader(tableLayout, "Course");

          */
                ////Add body  

                AddCellToHeader(tableLayout, "Course");
                AddCellToHeader(tableLayout, "Grade");

                foreach (var empee in students)
                {

                    /* AddCellToBody(tableLayout, emp.ID.ToString());
                     AddCellToBody(tableLayout, emp.FullName);
                     */

                    foreach (var empe in enrollments)
                    {

                        AddCellToBody(tableLayout, empe.Course.Title);
                        AddCellToBody(tableLayout, empe.Grade.ToString());





                    }

                }

            }

            return tableLayout;
        }


        // Make PDF matakuliah
        
        public ActionResult FileResultCreatePdfCoursesGrade()
        {
            MemoryStream workStream = new MemoryStream();
            StringBuilder status = new StringBuilder("");
            DateTime dTime = DateTime.Now;
            //file name to be created   
            string strPDFFileName = string.Format("SamplePdf" + dTime.ToString("yyyyMMdd") + "-" + ".pdf");
            Document doc = new Document();
            doc.SetMargins(10f, 10f, 10f, 10f);
            //Create PDF Table with 5 columns  
            PdfPTable tableLayout = new PdfPTable(2);
            doc.SetMargins(10f, 10f, 10f, 10f);
            //Create PDF Table  

            //file will created in this path  
            string strAttachment = Server.MapPath("~/Downloadss/" + strPDFFileName);


            PdfWriter.GetInstance(doc, workStream).CloseStream = false;
            doc.Open();

            //Add Content to PDF   
            doc.Add(Add_Content_To_PDF_courses(tableLayout));

            // Closing the document  
            doc.Close();

            byte[] byteInfo = workStream.ToArray();
            workStream.Write(byteInfo, 0, byteInfo.Length);
            workStream.Position = 0;


            return File(workStream, "application/pdf", strPDFFileName);

        }

        protected PdfPTable Add_Content_To_PDF_courses(PdfPTable tableLayout)
        {


            float[] headers = { 24, 45 }; //Header Widths  
            tableLayout.SetWidths(headers); //Set the pdf headers  
            tableLayout.WidthPercentage = 100; //Set the PDF File witdh percentage  
            tableLayout.HeaderRows = 1;
            //Add Title to the PDF file at the top  

            var students = from s in db.Students
                           select s;

            var enrollments = from e in db.Enrollments
                              select e;

            var courses = from c in db.Courses
                              select c;

            var instructors = from i in db.Instructors
                              select i;

       tableLayout.AddCell(new PdfPCell(new Phrase("Laporan Nilai Siswa untuk setiap Course", new Font(Font.FontFamily.HELVETICA, 8, 1, new iTextSharp.text.BaseColor(0, 0, 0))))
       {
           Colspan = 12,
           Border = 0,
           PaddingBottom = 4,
           HorizontalAlignment = Element.ALIGN_CENTER
       });

            foreach (var emp in courses)
            {

                foreach (var emm in instructors) {

                    AddCellToBody(tableLayout, "Course");
                    AddCellToBody(tableLayout, emp.Title);
                    AddCellToBody(tableLayout, "Instructor");
                    AddCellToBody(tableLayout, emm.FullName);
                }

                // tableLayout.AddCell(new PdfPCell(new Phrase(emp.Title)));
                /*ViewBag.NameSortParm = "Fullname";
                AddCellToBody(tableLayout, emp.FullName);

                ////Add header  
                /*  AddCellToHeader(tableLayout, "FullName");
                  AddCellToHeader(tableLayout, "Grade");
                  AddCellToHeader(tableLayout, "Course");

          */
                ////Add body  


                AddCellToHeader(tableLayout, "FullName");
                AddCellToHeader(tableLayout, "Grade");

                foreach (var empee in students)
                {

                    /* AddCellToBody(tableLayout, emp.ID.ToString());
                     AddCellToBody(tableLayout, emp.FullName);
                     */

                    foreach (var empe in enrollments)
                    {

                        AddCellToBody(tableLayout, empee.FullName);
                        AddCellToBody(tableLayout, empe.Grade.ToString());





                    }

                }

            }

            return tableLayout;
        }








        // Method to add single cell to the Header  
        private static void AddCellToHeader(PdfPTable tableLayout, string cellText)
        {

            tableLayout.AddCell(new PdfPCell(new Phrase(cellText, new Font(Font.FontFamily.HELVETICA, 8, 1, iTextSharp.text.BaseColor.YELLOW)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new iTextSharp.text.BaseColor(128, 0, 0)
            });
        }

        // Method to add single cell to the body  
        private static void AddCellToBody(PdfPTable tableLayout, string cellText)
        {
            tableLayout.AddCell(new PdfPCell(new Phrase(cellText, new Font(Font.FontFamily.HELVETICA, 8, 1, iTextSharp.text.BaseColor.BLACK)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new iTextSharp.text.BaseColor(255, 255, 255)
            });
        }




        // Make PDF Matakuliah



        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}