using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using MySql.Data.MySqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using System.Security.Cryptography;
namespace Transcripts_PDF_Generator
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }
       
        MySqlDataReader reader;
        MySqlCommand cmd;
        private void frmMain_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            generateTranscript();
        }
        private void studentDetails()
        {
           
        }
        private void generateTranscript() 
        {
            try
            {
                dbConnect dbConnect = new dbConnect();
                dbConnect.dbconnection();
               int count;
               cmd = new MySqlCommand("SELECT COUNT(admissionnumber)  FROM tblexamresults,tblstudents,tblschools,tbldepartments,tblcourses " +
                    "WHERE " +
                    "tblexamresults.studentid=tblstudents.studentid " +
                      "AND " +
                      "academicyear='23' AND tblstudents.courseid=tblcourses.courseid AND " +
                      "tbldepartments.departmentid=tblcourses.departmentid AND tblschools.schoolid=tbldepartments.schoolid", dbConnect.connection);
               
                count = Convert.ToInt32(cmd.ExecuteScalar().ToString());
                progressBar1.Maximum = count;
                progressBar1.Minimum = 0;
                progressBar1.Value = 0;

                dbConnect.connection.Close();
                dbConnect.dbconnection();
                 cmd = new MySqlCommand("SELECT *  FROM tblexamresults,tblstudents,tblschools,tbldepartments,tblcourses " +
                    "WHERE " +
                    "tblexamresults.studentid=tblstudents.studentid " +
                      "AND " +
                      "academicyear='23' AND tblstudents.courseid=tblcourses.courseid AND "+
                      "tbldepartments.departmentid=tblcourses.departmentid AND tblschools.schoolid=tbldepartments.schoolid", dbConnect.connection);
                
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    String studentid = reader["studentid"].ToString();
                    progressBar1.Value = progressBar1.Value + 1;
                    
                    string adminNo = reader["admissionnumber"].ToString().Replace("/", "-");
                    string EncryptedAdminNo = dbConnect.CalculateMD5Hash(adminNo.ToUpper());
                    EncryptedAdminNo = dbConnect.GetFullEncryption(EncryptedAdminNo.ToUpper()) +" "+studentid+" "+adminNo;
                               
                        string appRootDir = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName;

                       // FileStream fs = new FileStream(appRootDir + "/PDFs/" + adminNo + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None);
                        Document doc = new Document(PageSize.A4, 0f, 0f, 50f, 0f);
                        PdfWriter.GetInstance(doc, new FileStream(appRootDir + "/PDFs/" + EncryptedAdminNo + ".pdf", FileMode.Create));
                        

                            iTextSharp.text.Image tif = iTextSharp.text.Image.GetInstance(appRootDir + "/masai.png");

                            PdfPCell imageCell = new PdfPCell(tif);
                            imageCell.HorizontalAlignment = 1;
                            imageCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                            tif.ScalePercent(15f);
                            doc.Open();

                            PdfPTable table = new PdfPTable(3);
                            
                            BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);

                            iTextSharp.text.Font times = new iTextSharp.text.Font(bfTimes, 11);

                            var p = new Paragraph("P.O Box 861-20500", times);
                            p.Add(" ");
                            p.Add(" ");
                            p.Add(" ");
                            p.Add("Narok,Kenya");
                            PdfPCell website = new PdfPCell(p);
                            website.HorizontalAlignment = 2;

                            var p2 = new Paragraph("Tel: 0202685356/7", times);
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add(" ");
                            p2.Add("Email:registrar@nu.ac.ke");
                            PdfPCell email = new PdfPCell(p2);
                            email.HorizontalAlignment = 0;

                            iTextSharp.text.Font times2 = new iTextSharp.text.Font(bfTimes, 16, iTextSharp.text.Font.BOLD);
                            iTextSharp.text.Font times3 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.BOLD);
                            PdfPCell cell = new PdfPCell(new Phrase("MAASAI MARA UNIVERSITY", times2));
                            PdfPCell office = new PdfPCell(new Phrase("OFFICE OF THE DEAN SCHOOL OF SCIENCE", times3));
                            PdfPCell provisional = new PdfPCell(new Phrase("PROVISIONAL TRANSCRIPT", times3));
                            PdfPCell UNDERGRADUATE = new PdfPCell(new Phrase("UNDERGRADUATE", times3));
                            Paragraph line = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));
                            PdfPCell linecell = new PdfPCell(line);
                            cell.Colspan = 3;
                            cell.Border = iTextSharp.text.Rectangle.NO_BORDER;

                            office.Colspan = 3;
                            office.Border = iTextSharp.text.Rectangle.NO_BORDER;

                            provisional.Colspan = 3;
                            provisional.Border = iTextSharp.text.Rectangle.NO_BORDER;

                            UNDERGRADUATE.Colspan = 3;
                            UNDERGRADUATE.Border = iTextSharp.text.Rectangle.NO_BORDER;
                            linecell.Colspan = 3;
                            linecell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                            email.Border = iTextSharp.text.Rectangle.NO_BORDER;
                            website.Border = iTextSharp.text.Rectangle.NO_BORDER;
                            cell.HorizontalAlignment = 1; //0=Left, 1=Centre, 2=Right
                            office.HorizontalAlignment = 1;
                            provisional.HorizontalAlignment = 1;
                            UNDERGRADUATE.HorizontalAlignment = 1;
                            table.AddCell(email);

                            table.AddCell(imageCell);

                            table.AddCell(website);

                            table.AddCell(cell);
                            table.AddCell(office);
                            table.AddCell(provisional);
                            table.AddCell(UNDERGRADUATE);

                            table.AddCell(linecell);

                            doc.Add(table);
                            PdfPTable table2 = new PdfPTable(4);
                            table2.DefaultCell.Border = iTextSharp.text.Rectangle.NO_BORDER;
                           
                          
                            table2.AddCell(new Phrase("NAME: ", times3));

                            table2.AddCell(new Phrase(reader["surname"].ToString() + " " + reader["firstname"].ToString() + " " + reader["lastname"].ToString(), times3));

                            table2.AddCell(new Phrase("REG NO: ", times3));

                            table2.AddCell(new Phrase(reader["admissionnumber"].ToString(), times3));

                            table2.AddCell(new Phrase("SCHOOL OF: ", times3));

                            table2.AddCell(new Phrase(reader["schoolname"].ToString(), times3));

                            table2.AddCell(new Phrase("YEAR OF STUDY: ", times3));

                            table2.AddCell(new Phrase("THIRD YEAR: ", times3));

                            table2.AddCell(new Phrase("PROGRAMME: ", times3));

                            table2.AddCell(new Phrase(reader["coursename"].ToString(), times3));

                            table2.AddCell(new Phrase("ACADEMIC YEAR: ", times3));

                            table2.AddCell(new Phrase("2014/2015: ", times3));

                            doc.Add(table2);
                            PdfPTable table3 = new PdfPTable(8);
                            iTextSharp.text.Font times4 = new iTextSharp.text.Font(bfTimes, 10, iTextSharp.text.Font.NORMAL);
                            PdfPCell code = new PdfPCell(new Phrase("COURSE", times4));
                            code.Colspan = 4;
                            table3.AddCell(new Phrase("COURSE CODE", times3));
                            table3.AddCell(code);
                            table3.AddCell(new Phrase("UNITS", times3));
                            table3.AddCell(new Phrase("MARKS", times3));
                            table3.AddCell(new Phrase("GRADE", times3));
                           
                           MySqlConnection connection = new MySqlConnection(dbConnect.connectionString);
                           connection.Open();
                           cmd = new MySqlCommand("SELECT * FROM tblexamresults,tblunits WHERE tblexamresults.unitid=tblunits.unitid AND studentid='" + studentid + "' ", connection);
                            MySqlDataReader reader3;
                            
                        reader3 = cmd.ExecuteReader();
                         while(reader3.Read()){
                             PdfPCell coscode = new PdfPCell(new Phrase(reader3["unitname"].ToString(), times4));
                             coscode.Colspan = 4;
                              int totalMarks = Convert.ToInt32(reader3["totalmarks"].ToString());
                              String grade;
                              if (totalMarks>=70)
                              {
                                  grade = "A";
                              }
                              else if (totalMarks >= 60)
                              {
                                  grade = "B";
                              }
                              else if (totalMarks >= 50)
                              {
                                  grade = "C";
                              }
                              else if (totalMarks >= 40)
                              {
                                  grade = "D";
                              }
                              else {
                                  grade = "*";
                              }
                              table3.AddCell(new Phrase(reader3["unitcode"].ToString(), times4));
                              table3.AddCell(coscode);
                              table3.AddCell(new Phrase("3", times4));
                              table3.AddCell(new Phrase(reader3["totalmarks"].ToString(), times4));
                              table3.AddCell(new Phrase(grade, times4));
                          }

                            doc.Add(table3);
                            doc.Close();

                            connection.Close();
                    }
                
                dbConnect.connection.Close();
                
                MessageBox.Show("Pdf Created Successfully");
                progressBar1.Value = 0;
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
