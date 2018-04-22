using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    static class ProjectSummary
    {

        public static void Logo() {


            string title = @"
_________                        ________                     .__                                       __   
\_   ___ \  ___________   ____   \______ \   _______  __ ____ |  |   ____ ______   _____   ____   _____/  |_ 
/    \  \/ /  _ \_  __ \_/ __ \   |    |  \_/ __ \  \/ // __ \|  |  /  _ \\____ \ /     \_/ __ \ /    \   __\
\     \___(  <_> )  | \/\  ___/   |    `   \  ___/\   /\  ___/|  |_(  <_> )  |_> >  Y Y  \  ___/|   |  \  |  
 \______  /\____/|__|    \___  > /_______  /\___  >\_/  \___  >____/\____/|   __/|__|_|  /\___  >___|  /__|  
        \/                   \/          \/     \/          \/            |__|         \/     \/     \/      
          
                                                                 ";

            Console.WriteLine(title);
        }

        public static void ProjSummary1(ref Presentation p)
        {

            //formatting font text for table TextFrame

            PortionFormat portionFormat = new PortionFormat();
            portionFormat.FontHeight = 10;

            IMasterLayoutSlideCollection layoutSlides = p.Masters[0].LayoutSlides;
            ILayoutSlide layoutSlide =
                layoutSlides.GetByType(SlideLayoutType.TitleAndObject) ??
                layoutSlides.GetByType(SlideLayoutType.Title);
           
            p.Slides.AddEmptySlide(layoutSlide);
            
            #region Table Rows/Column
            double[] dblCols = { 20, 150, 50, 50, 150, 150, 50, 50 };
            double[] dblRows = { 30, 30, 30, 30, 30, 30, 30, 30, 30, 30 };
            #endregion

            #region project summary (1/3)
            ISlide sld = p.Slides[1];
            sld.Name = "Project Summary sheet 1";
           
            // Add an AutoShape of Rectangle type
            IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 35, 350, 50);
            ashp.FillFormat.FillType = FillType.NoFill;
            ashp.LineFormat.FillFormat.FillType = FillType.NoFill;

            // Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");
            
            // Accessing the text frame
            ITextFrame txtFrame = ashp.TextFrame;
           
            // Create the Paragraph object for text frame
            IParagraph para = txtFrame.Paragraphs[0];

            // Create Portion object for paragraph
            IPortion portion = para.Portions[0];
            portion.PortionFormat.FontHeight = 25;
            IPortionFormat pf = portion.PortionFormat;
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            portion.PortionFormat.FontBold = NullableBool.True;

            // Set Text
            portion.Text = "Projects Summary (1/3) – existing tenders (Priority 1)";

            
            //existing tenders 


            //Add table shape to slide

            ITable tb1 = sld.Shapes.AddTable(30, 100, dblCols, dblRows);
          
            for (int i = 0; i < tb1.Rows.Count; i++)
            {
                for (int j = 0; j < tb1.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb1[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb1[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb1[j, i].BorderTop.Width = 1;

                    tb1[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb1[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb1[j, i].BorderBottom.Width = 1;

                    tb1[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb1[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb1[j, i].BorderLeft.Width = 1;

                    tb1[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb1[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb1[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb1[j, i].FillFormat.FillType = FillType.Solid;
                    tb1[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first row
                    if (i == 0)
                    {
                        tb1[j, i].FillFormat.FillType = FillType.Solid;
                        tb1[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }

            tb1[0, 0].TextFrame.Text = "Nr";
            tb1[1, 0].TextFrame.Text = "Existing Tenders";
            tb1[2, 0].TextFrame.Text = "RAG current month";
            tb1[3, 0].TextFrame.Text = "RAG previous month";
            tb1[4, 0].TextFrame.Text = "Issue/Key Milestone";
            tb1[5, 0].TextFrame.Text = "Action agreed";
            tb1[6, 0].TextFrame.Text = "Project Manger";
            tb1[7, 0].TextFrame.Text = "Accountable";

            tb1.SetTextFormat(portionFormat);

            #endregion

            #region project summary(2/3)
            p.Slides.AddEmptySlide(layoutSlide);
            ISlide sld1 = p.Slides[2];
            
            sld1.Name = "Project Summary sheet 2";

            // Add an AutoShape of Rectangle type
            ashp = sld1.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 35, 400, 50);
            ashp.FillFormat.FillType = FillType.NoFill;
            ashp.LineFormat.FillFormat.FillType = FillType.NoFill;
            //// Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");

            // Accessing the text frame
             txtFrame = ashp.TextFrame;

           // Create the Paragraph object for text frame

             para = txtFrame.Paragraphs[0];

           // Create Portion object for paragraph
            portion = para.Portions[0];
            portion.PortionFormat.FontHeight = 25;

            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            portion.PortionFormat.FontBold = NullableBool.True;

            // Set Text
            portion.Text = "Projects Summary (2/3) – new tenders";

            //Add table shape to slide

            ITable tb2 = sld1.Shapes.AddTable(30, 100, dblCols, dblRows);

            for (int i = 0; i < tb2.Rows.Count; i++)
            {
                for (int j = 0; j < tb2.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb2[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb2[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb2[j, i].BorderTop.Width = 1;

                    tb2[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb2[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb2[j, i].BorderBottom.Width = 1;

                    tb2[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb2[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb2[j, i].BorderLeft.Width = 1;

                    tb2[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb2[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb2[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb2[j, i].FillFormat.FillType = FillType.Solid;
                    tb2[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first row
                    if (i == 0)
                    {
                        tb2[j, i].FillFormat.FillType = FillType.Solid;
                        tb2[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }

            tb2[0, 0].TextFrame.Text = "Nr";
            tb2[1, 0].TextFrame.Text = "Existing Tenders";
            tb2[2, 0].TextFrame.Text = "RAG current month";
            tb2[3, 0].TextFrame.Text = "RAG previous month";
            tb2[4, 0].TextFrame.Text = "Issue/Key Milestone";
            tb2[5, 0].TextFrame.Text = "Action agreed";
            tb2[6, 0].TextFrame.Text = "Project Manger";
            tb2[7, 0].TextFrame.Text = "Accountable";

            tb2.SetTextFormat(portionFormat);
            #endregion

            #region project summary(3/3)

            p.Slides.AddEmptySlide(layoutSlide);
            ISlide sld2 = p.Slides[3];

            sld2.Name = "Project Summary sheet 3";

            // Add an AutoShape of Rectangle type
            ashp = sld2.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 35, 400, 50);
            ashp.FillFormat.FillType = FillType.NoFill;
            ashp.LineFormat.FillFormat.FillType = FillType.NoFill;
            //// Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");

            // Accessing the text frame
            txtFrame = ashp.TextFrame;

            // Create the Paragraph object for text frame
            para = txtFrame.Paragraphs[0];

            // Create Portion object for paragraph
            portion = para.Portions[0];
            portion.PortionFormat.FontHeight = 25;

            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            portion.PortionFormat.FontBold = NullableBool.True;

            // Set Text
            portion.Text = "Projects Summary (3/3) – M&A existing and new countries";

            //Add table shape to slide

            ITable tb3 = sld2.Shapes.AddTable(30, 100, dblCols, dblRows);

            for (int i = 0; i < tb3.Rows.Count; i++)
            {
                for (int j = 0; j < tb3.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb3[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb3[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3[j, i].BorderTop.Width = 1;

                    tb3[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb3[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3[j, i].BorderBottom.Width = 1;

                    tb3[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb3[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3[j, i].BorderLeft.Width = 1;

                    tb3[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb3[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb3[j, i].FillFormat.FillType = FillType.Solid;
                    tb3[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first row
                    if (i == 0)
                    {
                        tb3[j, i].FillFormat.FillType = FillType.Solid;
                        tb3[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }

            tb3[0, 0].TextFrame.Text = "Nr";
            tb3[1, 0].TextFrame.Text = "Growth Opportunities (Tenders): Priority 1";
            tb3[2, 0].TextFrame.Text = "RAG current month";
            tb3[3, 0].TextFrame.Text = "RAG previous month";
            tb3[4, 0].TextFrame.Text = "Issue/Key Milestone";
            tb3[5, 0].TextFrame.Text = "Action agreed";
            tb3[6, 0].TextFrame.Text = "Project Manger";
            tb3[7, 0].TextFrame.Text = "Accountable";
            tb3.SetTextFormat(portionFormat);
            #endregion
        }


        public static void OpportunitiesSchedule(ref Presentation p) {


            PortionFormat portionFormat = new PortionFormat();
            portionFormat.FontHeight = 10;

            IMasterLayoutSlideCollection layoutSlides = p.Masters[0].LayoutSlides;
            ILayoutSlide layoutSlide =
                layoutSlides.GetByType(SlideLayoutType.TitleAndObject) ??
                layoutSlides.GetByType(SlideLayoutType.Title);
            //add new slide in presentation
            p.Slides.AddEmptySlide(layoutSlide);
            
            #region Table Rows/Column
            double[] dblCols = { 10, 20, 30, 100, 20, 40, 40,40,50,30,30,30,30,30,30,30,30,40,70 };
            double[] dblRows = { 20, 20, 20, 20, 20, 20, 20, 20, 20, 20 };
            #endregion


            #region Opportunities Schedule – Existing Tenders by Priority
            ISlide sld = p.Slides[4];
            sld.Name = "Opportunities Schedule  sheet 1";

            // Add an AutoShape of Rectangle type
            IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 35, 350, 50);

            ashp.FillFormat.FillType = FillType.NoFill;
            ashp.LineFormat.FillFormat.FillType = FillType.NoFill;
            // Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");

            // Accessing the text frame
            ITextFrame txtFrame = ashp.TextFrame;

            // Create the Paragraph object for text frame
            IParagraph para = txtFrame.Paragraphs[0];

            // Create Portion object for paragraph
            IPortion portion = para.Portions[0];
            portion.PortionFormat.FontHeight = 25;
           
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            portion.PortionFormat.FontBold = NullableBool.True;

            // Set Text
            portion.Text = "Opportunities Schedule – Existing Tenders by Priority";

            //Add table shape to slide

            ITable tb1 = sld.Shapes.AddTable(5, 100, dblCols, dblRows);

            for (int i = 0; i < tb1.Rows.Count; i++)
            {
                for (int j = 0; j < tb1.Rows[i].Count; j++)
                {
                    if (i == 9 && (j < 8 || j > 17))
                    {
                        tb1[j, i].FillFormat.FillType = FillType.NoFill;
                    }
                    else
                    {
                        //adding border to each cell of the table
                        tb1[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                        tb1[j, i].BorderTop.Width = 1;

                        tb1[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                        tb1[j, i].BorderBottom.Width = 1;

                        tb1[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                        tb1[j, i].BorderLeft.Width = 1;

                        tb1[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                        tb1[j, i].BorderRight.Width = 1;
                        //cell with white background
                        tb1[j, i].FillFormat.FillType = FillType.Solid;
                        tb1[j, i].FillFormat.SolidFillColor.Color = Color.White;
                    }
                    //Grey color for first row
                    if (i == 0)
                    {
                       
                        tb1[j, i].FillFormat.FillType = FillType.Solid;
                        tb1[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                       

                       
                    }

                   
                }
            }

            for (int j = 0; j < tb1.Rows[0].Count; j++) {
                //Merge first and second row in header for specific column
                if (j < 8 || j > 17)
                    tb1.MergeCells(tb1[j, 0], tb1[j, 1], false);
                // merge three columns in first row
                if (j == 9||j==12||j==15) {
                    tb1.MergeCells(tb1[j, 0], tb1[j+1, 0], false);
                    tb1.MergeCells(tb1[j, 0], tb1[j + 2, 0], false);
                }
            }

            tb1[0, 0].TextFrame.Text = "#";
            tb1[1, 0].TextFrame.Text = "BU";
            tb1[2, 0].TextFrame.Text = "Priority";
            tb1[3, 0].TextFrame.Text = "Existing Tenders";
            tb1[4, 0].TextFrame.Text = "Stage";
            tb1[5, 0].TextFrame.Text = "Contract Date";
            tb1[6, 0].TextFrame.Text = "Directors Approval";
            tb1[7, 0].TextFrame.Text = "Team Yrs.";
            tb1[8, 0].TextFrame.Text = "Category";
            tb1[9, 0].TextFrame.Text = " ";
            tb1[10, 0].TextFrame.Text = "";
            tb1[11, 0].TextFrame.Text = "Revenue (€m)";
            tb1[12, 0].TextFrame.Text = "";
            tb1[13, 0].TextFrame.Text = "";
            tb1[14, 0].TextFrame.Text = "EBIT (€m)";
            tb1[15, 0].TextFrame.Text = "";
            tb1[16, 0].TextFrame.Text = "";
            tb1[17, 0].TextFrame.Text = "CAPEX (€m)";
            tb1[18, 0].TextFrame.Text = "EBIT/CAPEX";


            tb1[8, 9].TextFrame.Text = "Total MTP";
            //txtFrame = tb1[8, 9].TextFrame;

            tb1.SetTextFormat(portionFormat);
            StageInfo(ref sld);
            #endregion

            #region "Opportunities Schedule – New Tenders by priority"

            p.Slides.AddEmptySlide(layoutSlide);
            ISlide sld1 = p.Slides[5];
            #region Table Rows/Column
            double[] dblCols1 = { 10, 20, 30, 100, 20, 40, 40, 40, 50, 30, 30, 30, 30, 30, 30, 30, 30, 40, 70 };
            double[] dblRows1 = { 20, 20, 20, 20, 20, 20, 20, 20, 20, 20 };
            #endregion

            IAutoShape ashp1 = sld1.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 35, 350, 50);

            ashp1.FillFormat.FillType = FillType.NoFill;
            ashp1.LineFormat.FillFormat.FillType = FillType.NoFill;
           
            ashp1.AddTextFrame(" ");
           
            // Accessing the text frame
            ITextFrame txtFrame1 = ashp1.TextFrame;

            // Create the Paragraph object for text frame
            IParagraph para1 = txtFrame1.Paragraphs[0];

            // Create Portion object for paragraph
            IPortion portion1 = para1.Portions[0];
            portion1.PortionFormat.FontHeight = 25;

            portion1.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion1.PortionFormat.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            portion1.PortionFormat.FontBold = NullableBool.True;
         
            // Set Text
            portion1.Text = "Opportunities Schedule – New Tenders by Priority";

            //Add table shape to slide

            ITable tb2 = sld1.Shapes.AddTable(5, 100, dblCols, dblRows);
            for (int i = 0; i < tb2.Rows.Count; i++)
            {
                for (int j = 0; j < tb2.Rows[i].Count; j++)
                {
                    if ((i == 9 || i == 4) && (j < 8 || j > 17))
                    {
                        tb2[j, i].FillFormat.FillType = FillType.NoFill;
                    }
                    //else if (i == 4 && (j < 8 || j > 17))
                    //{
                    //    tb2[j, i].FillFormat.FillType = FillType.NoFill;
                    //}
                    else
                    {
                        //adding border to each cell of the table
                        tb2[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                        tb2[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                        tb2[j, i].BorderTop.Width = 1;

                        tb2[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                        tb2[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                        tb2[j, i].BorderBottom.Width = 1;

                        tb2[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                        tb2[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                        tb2[j, i].BorderLeft.Width = 1;

                        tb2[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                        tb2[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                        tb2[j, i].BorderRight.Width = 1;
                        //cell with white background
                        tb2[j, i].FillFormat.FillType = FillType.Solid;
                        tb2[j, i].FillFormat.SolidFillColor.Color = Color.White;
                    }
                    //Grey color for first row
                    if (i == 0)
                    {

                        tb2[j, i].FillFormat.FillType = FillType.Solid;
                        tb2[j, i].FillFormat.SolidFillColor.Color = Color.Gray;



                    }


                }
            }

            for (int j = 0; j < tb2.Rows[0].Count; j++)
            {
                //Merge first and second row in header for specific column
                if (j < 8 || j > 17)
                    tb2.MergeCells(tb2[j, 0], tb2[j, 1], false);
                // merge three columns in first row
                if (j == 9 || j == 12 || j == 15)
                {
                    tb2.MergeCells(tb2[j, 0], tb2[j + 1, 0], false);
                    tb2.MergeCells(tb2[j, 0], tb2[j + 2, 0], false);
                }
            }
            tb2[8, 4].TextFrame.Text = "TOTAL MTP";
            tb2[0, 0].TextFrame.Text = "#";
            tb2[1, 0].TextFrame.Text = "BU";
            tb2[2, 0].TextFrame.Text = "Priority";
            tb2[3, 0].TextFrame.Text = "New Tenders";
            tb2[4, 0].TextFrame.Text = "Stage";
            tb2[5, 0].TextFrame.Text = "Contract Date";
            tb2[6, 0].TextFrame.Text = "Directors Approval";
            tb2[7, 0].TextFrame.Text = "Team Yrs.";
            tb2[8, 0].TextFrame.Text = "Category";
            tb2[9, 0].TextFrame.Text = " ";
            tb2[10, 0].TextFrame.Text = "";
            tb2[11, 0].TextFrame.Text = "Revenue (€m)";
            tb2[12, 0].TextFrame.Text = "";
            tb2[13, 0].TextFrame.Text = "";
            tb2[14, 0].TextFrame.Text = "EBIT (€m)";
            tb2[15, 0].TextFrame.Text = "";
            tb2[16, 0].TextFrame.Text = "";
            tb2[17, 0].TextFrame.Text = "CAPEX (€m)";
            tb2[18, 0].TextFrame.Text = "EBIT/CAPEX";


            tb2[8, 9].TextFrame.Text = "Total MTP";
            //txtFrame = tb2[8, 9].TextFrame;

            tb2.SetTextFormat(portionFormat);
            StageInfo(ref sld1);
            #endregion




        }


        // Generic functions for Slides

        #region Stage Info
        private static void StageInfo(ref ISlide sld)
        {
            

            IAutoShape ashp1 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 20, 300, 200, 300);


            ashp1.FillFormat.FillType = FillType.NoFill;
            ashp1.LineFormat.FillFormat.FillType = FillType.NoFill;
            ITextFrame txtFrame1 = ashp1.TextFrame;
            txtFrame1.TextFrameFormat.AutofitType = TextAutofitType.Shape;

            IParagraph para1 = txtFrame1.Paragraphs[0];

            para1.Text = "\n Stage 1 = Monitoring \n Stage 2 = ITT Assessment \n Stage 3 = Bid Preparation \n Stage 4 = Bid Submission \n Stage 5 = Mobilisation \n Stage 6 = Post tender review";
            para1.ParagraphFormat.Alignment = TextAlignment.Left;
            // Create Portion object for paragraph
            IPortion portion1 = para1.Portions[0];
            portion1.PortionFormat.FontHeight = 15;
            portion1.PortionFormat.FontItalic = NullableBool.True;
            portion1.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion1.PortionFormat.FillFormat.SolidFillColor.Color = Color.Black;

           }


        #endregion

    }
}
