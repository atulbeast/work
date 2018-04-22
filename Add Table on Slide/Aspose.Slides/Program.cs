using Aspose.Slides.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    class Program
    {
        //for authorize purpose
        public static async Task MainAsync(string[] args)
        {
            using (var client = new HttpClient())
            {
                //client.DefaultRequestHeaders.Accept.Clear();
                //client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                //client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "For Security Reason AccessToken Hidden");
                //Console.WriteLine("Accessing  Api Request Url");
                //var response = await client.GetAsync("http://localhost:61718/api/Project/Get?region=East%20Europe&countryCode=TR");
               
                //var results = default(string);
                //if (response.IsSuccessStatusCode)
                //{
                //    var contentResult = response.Content.ReadAsStringAsync();
                //    contentResult.Wait();
                //    results = contentResult.Result;
                //}
                //else
                //{
                //    results = response.StatusCode.ToString();
                //}
                //Console.BackgroundColor = ConsoleColor.Black;
            }
            Console.ReadKey();
        }

        static void Main(string[] args)
        {
            MainAsync(args).Wait();
            string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName.ToString(); // Directory Path
            string JsonData = File.ReadAllText(Path.GetFullPath(path) + "\\Json\\project.json"); //for temporary usage ..for production mainasync method will be used for fetching data
             ProjectModel pmodel = new ProjectModel();
            pmodel= JsonConvert.DeserializeObject<ProjectModel>(JsonData);
            //Instantiate Prsentation class that represents the PPTX
            Presentation pres = new Presentation();
            //Get the first slide
           
            string currentDir = Path.GetFullPath(path) + "\\pptFolder\\";
            ISlide sld = pres.Slides[0];
            sld.Name = "Existing Tender";
            //formatting Text
            PortionFormat portionFormat = new PortionFormat();
            portionFormat.FontHeight = 10;
            ProjectSummary.Logo();

            #region Border
            IAutoShape borderLines = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 3, 3, 715, 535);
            borderLines.FillFormat.FillType = FillType.NoFill;
            borderLines.FillFormat.SolidFillColor.Color = Color.Black;
           #endregion

            #region Background

            //Add some text
            IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 350, 20);
            //ashp.FillFormat.FillType = FillType.NoFill;
            //ashp.LineFormat.FillFormat.FillType = FillType.NoFill;
            // Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");
            ashp.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            // Accessing the text frame
            ITextFrame txtFrame = ashp.TextFrame;
            // Create the Paragraph object for text frame
            IParagraph para = txtFrame.Paragraphs[0];
            // Create Portion object for paragraph
            IPortion portion = para.Portions[0];
            // Set Text
            portion.Text = "Background";
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.White;
            portion.PortionFormat.FontBold = NullableBool.True;


            #region Table 1 
            double[] dblCols1 = { 60, 100, 190 };
            double[] dblRows1 = { 5, 5, 5, 5, 5, 5, 5, 5, 5, 5 };

            ITable tb1 = sld.Shapes.AddTable(10, 40, dblCols1, dblRows1);

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

                    //Grey color for first column
                    if (j == 0)
                    {
                        tb1[j, i].FillFormat.FillType = FillType.Solid;
                        tb1[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }

                    //for 7th to last custom row 
                    if (i >= 7)
                    {

                        tb1[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.White;
                        tb1[j, i].BorderLeft.Width = 1;

                        tb1[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.White;
                        tb1[j, i].BorderBottom.Width = 1;

                        tb1[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.White;
                        tb1[j, i].BorderRight.Width = 1;
                        tb1[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                        tb1[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.White;
                        tb1[j, i].BorderTop.Width = 1;
                        if (i == 7)
                        {
                            tb1[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                            tb1[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                            tb1[j, i].BorderTop.Width = 1;
                        }
                        if (j == 2)
                        {
                            tb1[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                            tb1[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                            tb1[j, i].BorderLeft.Width = 1;

                        }
                        if (j == 2)
                        {
                            tb1[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                            tb1[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                            tb1[j, i].BorderLeft.Width = 1;

                        }
                        tb1[j, i].FillFormat.FillType = FillType.Solid;
                        tb1[j, i].FillFormat.SolidFillColor.Color = Color.White;
                    }

                }

            }

            tb1.SetTextFormat(portionFormat);

            Bitmap image = new Bitmap(currentDir + "imagelogo.jpg");
            // Create an IPPImage object using the bitmap object
            IPPImage imgx1 = pres.Images.AddImage(image);
          
            // Add image to column 2 table cell
            tb1[2, 0].FillFormat.FillType = FillType.Picture;
            tb1[2, 0].FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
            tb1[2, 0].FillFormat.PictureFillFormat.Picture.Image = imgx1;


            tb1.MergeCells(tb1[2, 0], tb1[2, 1], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 2], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 3], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 4], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 5], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 6], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 7], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 8], false);
            tb1.MergeCells(tb1[2, 0], tb1[2, 9], false);

            #endregion
            //column Names:
            tb1[0, 0].TextFrame.Text = "ProjectName";
            tb1[0, 1].TextFrame.Text = "Country";
            tb1[0, 2].TextFrame.Text = "Region";
            tb1[0, 3].TextFrame.Text = "Mode";
            tb1[0, 4].TextFrame.Text = "Project Type";
            tb1[0, 5].TextFrame.Text = "Project Id";
            tb1[0, 6].TextFrame.Text = "Updated Date";

            //inserting values from project

            tb1[1, 0].TextFrame.Text = pmodel.ProjectName;
            tb1[1, 1].TextFrame.Text = pmodel.Country;
            tb1[1, 2].TextFrame.Text = pmodel.Region;
            tb1[1, 3].TextFrame.Text = Enum.GetName(typeof(EnumValue.Mode), pmodel.Mode);
            tb1[1, 4].TextFrame.Text = Enum.GetName(typeof(EnumValue.TenderStatus), pmodel.DealStatus); 
            tb1[1, 5].TextFrame.Text = pmodel.ProjectId;
            tb1[1, 6].TextFrame.Text = "7/09/2017";  



            #region Table 2

            double[] coltbl2 = { 350 };
            double[] rowtbl2 = { 20, 50, 20, 50 };
            ITable tb2 = sld.Shapes.AddTable(10, 230, coltbl2, rowtbl2);

            for (int i = 0; i < tb2.Rows.Count; i++)
            {

                for (int j = 0; j < tb2.Rows[i].Count; j++)
                {
                    //border for each cell
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

                    tb2[j, i].FillFormat.FillType = FillType.Solid;
                    if (i % 2 == 0)
                        tb2[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    else
                        tb2[j, i].FillFormat.SolidFillColor.Color = Color.White;

                }
            }

            //Text Entry
            tb2[0, 0].TextFrame.Text = "Description";
            tb2[0, 1].TextFrame.Text = pmodel.Description;
            tb2[0, 2].TextFrame.Text = "Strategic Rationale";
            tb2[0, 3].TextFrame.Text = pmodel.StrategicRationale;
            tb2.SetTextFormat(portionFormat);

            #endregion

            #region Table 3
            // Table 3a

            double[] coltbl3a = { 40, 40, 40, 40 };
            double[] rowtbl3a = { 10, 10, 10, 10, 10 };
            ITable tb3a = sld.Shapes.AddTable(10, 400, coltbl3a, rowtbl3a);

            portionFormat.FontHeight = 10;
            tb3a.SetTextFormat(portionFormat);
            for (int i = 0; i < tb3a.Rows.Count; i++)
            {
                for (int j = 0; j < tb3a.Rows[i].Count; j++)
                {
                    tb3a[j, i].FillFormat.FillType = FillType.Solid;
                    if (i == 0 && j == 0)
                    {
                        tb3a[j, i].FillFormat.FillType = FillType.NoFill;
                    }
                    else
                    {
                        if ((i == 0 && j > 0) || (j == 0 && i > 0))
                        {
                            tb3a[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                        }
                        else
                            tb3a[j, i].FillFormat.SolidFillColor.Color = Color.White;
                        //border for each cell
                        tb3a[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                        tb3a[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                        tb3a[j, i].BorderTop.Width = 1;

                        tb3a[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                        tb3a[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                        tb3a[j, i].BorderBottom.Width = 1;

                        tb3a[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                        tb3a[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                        tb3a[j, i].BorderLeft.Width = 1;

                        tb3a[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                        tb3a[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                        tb3a[j, i].BorderRight.Width = 1;


                    }


                }
            }
            //column names:
            tb3a[1, 0].TextFrame.Text = "2017";
            tb3a[2, 0].TextFrame.Text = "2018";
            tb3a[3, 0].TextFrame.Text = "2019";
            tb3a[0, 1].TextFrame.Text = "Revenue";
            tb3a[0, 2].TextFrame.Text = "EBIT";
            tb3a[0, 3].TextFrame.Text = "Capex";
            tb3a[0, 4].TextFrame.Text = "MTP";

            tb3a[1, 1].TextFrame.Text = pmodel.RevenueX.ToString();
            tb3a[2, 1].TextFrame.Text = pmodel.RevenueX1.ToString();
            tb3a[3, 1].TextFrame.Text = pmodel.RevenueX2.ToString();

            tb3a[1, 2].TextFrame.Text = pmodel.EbitX.ToString();
            tb3a[2, 2].TextFrame.Text = pmodel.EbitX1.ToString();
            tb3a[3, 2].TextFrame.Text = pmodel.EbitX2.ToString();

            tb3a[1, 3].TextFrame.Text = pmodel.CAPEXXMEuro.ToString();
            tb3a[2,3].TextFrame.Text = pmodel.CAPEX1MEuro.ToString();
            tb3a[3, 3].TextFrame.Text = pmodel.CAPEX2MEuro.ToString();
            // Table 3b
            double[] coltbl3b = { 100, 50 };
            double[] rowtbl3b = { 10, 10, 10 };
            ITable tb3b = sld.Shapes.AddTable(200, 400, coltbl3b, rowtbl3b);

            portionFormat.FontHeight = 10;
            tb3b.SetTextFormat(portionFormat);
            for (int i = 0; i < tb3b.Rows.Count; i++)
            {
                for (int j = 0; j < tb3b.Rows[i].Count; j++)
                {
                    tb3b[j, i].FillFormat.FillType = FillType.Solid;

                    if (j == 0)
                    {
                        tb3b[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                    else
                        tb3b[j, i].FillFormat.SolidFillColor.Color = Color.White;
                    //border for each cell
                    tb3b[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb3b[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3b[j, i].BorderTop.Width = 1;

                    tb3b[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb3b[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3b[j, i].BorderBottom.Width = 1;

                    tb3b[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb3b[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3b[j, i].BorderLeft.Width = 1;

                    tb3b[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb3b[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb3b[j, i].BorderRight.Width = 1;


                }
            }


            //columns name
            tb3b[0, 0].TextFrame.Text = "EBIT";
            tb3b[0, 1].TextFrame.Text = "Number of vehicles";
            tb3b[0, 2].TextFrame.Text="Contract Length";

            tb3b[0, 0].TextFrame.Text = pmodel.New_or_existingwork.ToString();
            tb3b[0, 1].TextFrame.Text = pmodel.NumberOfVehicles.ToString();
            tb3b[0, 2].TextFrame.Text = pmodel.CoreContractLength.ToString();



            #endregion

            #endregion

            #region Status

            //Add some text
            IAutoShape ashp1 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 400, 10, 300, 20);
           
            // Add TextFrame to the Rectangle
            ashp1.AddTextFrame("");
            ashp1.FillFormat.SolidFillColor.Color = Color.DarkCyan;
            // Accessing the text frame
            ITextFrame txtFrame1 = ashp1.TextFrame;

          
            // Create the Paragraph object for text frame
            IParagraph para1 = txtFrame1.Paragraphs[0];
            // Create Portion object for paragraph
            IPortion portion1 = para1.Portions[0];
            // Set Text
            portion1.Text = "Status";
            portion1.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion1.PortionFormat.FillFormat.SolidFillColor.Color = Color.White;
            portion1.PortionFormat.FontBold = NullableBool.True;

            #region Table 4
            //Define columns with widths and rows with heights
            //double[] dblCols = { 50, 50, 50 };
            double[] dblCols = { 80, 100, 60, 60 };
            double[] dblRows = { 20, 20, 20, 20, 40, 30 };

            //Add table shape to slide

            ITable tb4 = sld.Shapes.AddTable(400, 40, dblCols, dblRows);
            tb4.SetTextFormat(portionFormat);
            for (int i = 0; i < tb4.Rows.Count; i++)
            {
                for (int j = 0; j < tb4.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb4[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb4[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb4[j, i].BorderTop.Width = 1;

                    tb4[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb4[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb4[j, i].BorderBottom.Width = 1;

                    tb4[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb4[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb4[j, i].BorderLeft.Width = 1;

                    tb4[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb4[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb4[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb4[j, i].FillFormat.FillType = FillType.Solid;
                    tb4[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first column
                    if (j == 0)
                    {
                        tb4[j, i].FillFormat.FillType = FillType.Solid;
                        tb4[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }

            //color for cell
            //caution: color it before split otherwise left portion of split cell will be coloured
            tb4[3, 5].FillFormat.FillType = FillType.Solid;
            tb4[3, 5].FillFormat.SolidFillColor.Color = Color.ForestGreen;
            tb4.SetTextFormat(portionFormat);


            //Set border format for each cell
            //Merge cells 1 & 2 of row 1
            tb4.MergeCells(tb4[1, 4], tb4[2, 4], false);
            tb4.MergeCells(tb4[2, 4], tb4[3, 4], false);
            tb4.MergeCells(tb4[1, 5], tb4[2, 5], false);
            tb4[3, 5].SplitByWidth(tb4[3, 5].Width / 2);
            tb4.MergeCells(tb4[2, 5], tb4[3, 5],false);
            //  tb4[1, 4].TextFrame.Text = "";



            //column names: 

            tb4[0, 0].TextFrame.Text = "Division Responsible";
            tb4[0, 1].TextFrame.Text = "Executive Board Member";
            tb4[0, 2].TextFrame.Text = "Country Manager";
            tb4[0, 3].TextFrame.Text = "Project Manager";
            tb4[0, 4].TextFrame.Text = "Team Members(and role) OR requirements";
            tb4[0, 5].TextFrame.Text = "Uncovered Team Resources";
            tb4[2, 0].TextFrame.Text = "Contract Type";
            tb4[2, 1].TextFrame.Text = "Project Stage";
            tb4[2, 2].TextFrame.Text = "Type";
            tb4[2, 3].TextFrame.Text = "Probability";

            //project values:
            tb4[1, 0].TextFrame.Text = pmodel.AdditionalTeamMembers;
            tb4[1, 1].TextFrame.Text = pmodel.ExecutiveBoardMember;
            tb4[1, 2].TextFrame.Text = pmodel.CountryManager;
            tb4[1, 3].TextFrame.Text = pmodel.ProjectManager;
            tb4[1, 4].TextFrame.Text = pmodel.UncoveredTeamResources;
            tb4[2, 0].TextFrame.Text = Enum.GetName(typeof(EnumValue.ContractType), pmodel.ContractType);
            tb4[2, 1].TextFrame.Text = Enum.GetName(typeof(EnumValue.ContractType), pmodel.ProjectStageTender);
            tb4[2, 2].TextFrame.Text = "A(>500)";
            tb4[2, 3].TextFrame.Text = "60%";

            #endregion

            #region Table 5

            double[] dblCols5 = { 80, 160, 30, 30 };
            double[] dblRows5 = { 10, 10, 10, 10, 10, 10,10,10,10 };
            ITable tb5 = sld.Shapes.AddTable(400,200, dblCols5, dblRows5);
            for (int i = 0; i < tb5.Rows.Count; i++)
            {
                for (int j = 0; j < tb5.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb5[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb5[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb5[j, i].BorderTop.Width = 1;

                    tb5[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb5[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb5[j, i].BorderBottom.Width = 1;

                    tb5[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb5[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb5[j, i].BorderLeft.Width = 1;

                    tb5[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb5[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb5[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb5[j, i].FillFormat.FillType = FillType.Solid;
                    tb5[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first column
                    if (i < 2||j==3)
                    {
                        tb5[j, i].FillFormat.FillType = FillType.Solid;
                        tb5[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }
            tb5.SetTextFormat(portionFormat);

            //column names: 
            tb5[0, 0].TextFrame.Text = "KeyMilestone & Action";
            tb5[0, 1].TextFrame.Text = "Date";
            tb5[1, 1].TextFrame.Text = "Task/Event";
            tb5[2, 1].TextFrame.Text = "Resp";
            tb5[3, 1].TextFrame.Text = "Status";
            for (int i = 2; i <= pmodel.KeyMileStoneAndAction.Count+1; i++)
            {

                
                tb5[0, i].TextFrame.Text = pmodel.KeyMileStoneAndAction[i - 2].Date;
                tb5[1, i].TextFrame.Text = pmodel.KeyMileStoneAndAction[i - 2].TaskEvent;
                tb5[2, i].TextFrame.Text = pmodel.KeyMileStoneAndAction[i - 2].Resp;
                tb5[3, i].TextFrame.Text = pmodel.KeyMileStoneAndAction[i - 2].Status;

            }

            #endregion

            #region Table 6
            double[] dblCols6 = { 300 };
            double[] dblRows6 = {20,40};

            ITable tb6 = sld.Shapes.AddTable(400, 380, dblCols6, dblRows6);
            tb6.SetTextFormat(portionFormat);
            for (int i = 0; i < tb6.Rows.Count; i++)
            {
                for (int j = 0; j < tb6.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb6[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb6[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb6[j, i].BorderTop.Width = 1;

                    tb6[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb6[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb6[j, i].BorderBottom.Width = 1;

                    tb6[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb6[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb6[j, i].BorderLeft.Width = 1;

                    tb6[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb6[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb6[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb6[j, i].FillFormat.FillType = FillType.Solid;
                    tb6[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first column
                    if (i==0 && j==0)
                    {
                        tb6[j, i].FillFormat.FillType = FillType.Solid;
                        tb6[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }

            #endregion

            tb6[0, 0].TextFrame.Text = "Status Description";
            tb6[0, 1].TextFrame.Text = pmodel.StatusDescription;


            #region Table 7
            double[] dblCols7 = { 240,30,30 };
            double[] dblRows7 = { 20,20,20,20 };

            ITable tb7 = sld.Shapes.AddTable(400, 445, dblCols7, dblRows7);
            tb7.SetTextFormat(portionFormat);
            for (int i = 0; i < tb7.Rows.Count; i++)
            {
                for (int j = 0; j < tb7.Rows[i].Count; j++)
                {
                    //adding border to each cell of the table
                    tb7[j, i].BorderTop.FillFormat.FillType = FillType.Solid;
                    tb7[j, i].BorderTop.FillFormat.SolidFillColor.Color = Color.Black;
                    tb7[j, i].BorderTop.Width = 1;

                    tb7[j, i].BorderBottom.FillFormat.FillType = FillType.Solid;
                    tb7[j, i].BorderBottom.FillFormat.SolidFillColor.Color = Color.Black;
                    tb7[j, i].BorderBottom.Width = 1;

                    tb7[j, i].BorderLeft.FillFormat.FillType = FillType.Solid;
                    tb7[j, i].BorderLeft.FillFormat.SolidFillColor.Color = Color.Black;
                    tb7[j, i].BorderLeft.Width = 1;

                    tb7[j, i].BorderRight.FillFormat.FillType = FillType.Solid;
                    tb7[j, i].BorderRight.FillFormat.SolidFillColor.Color = Color.Black;
                    tb7[j, i].BorderRight.Width = 1;
                    //cell with white background
                    tb7[j, i].FillFormat.FillType = FillType.Solid;
                    tb7[j, i].FillFormat.SolidFillColor.Color = Color.White;

                    //Grey color for first column
                    if (i == 0)
                    {
                        tb7[j, i].FillFormat.FillType = FillType.Solid;
                        tb7[j, i].FillFormat.SolidFillColor.Color = Color.Gray;
                    }
                }
            }

            #endregion
            tb7[0, 0].TextFrame.Text = "Issue For Discussion";
            for (int i = 1; i <= pmodel.Issues.Count; i++)
            {
                tb7[0, i].TextFrame.Text = pmodel.Issues[i-1].Issues;
                tb7[1, i].TextFrame.Text = pmodel.Issues[i-1].Owner;
                tb7[2, i].TextFrame.Text = pmodel.Issues[i-1].Date;
            }
         

            #endregion

            #region projectsummary

            ProjectSummary.ProjSummary1(ref pres);

            #endregion

            #region Opportunities Schedule

            ProjectSummary.OpportunitiesSchedule(ref pres);


            #endregion


            Console.Write("ppt in progress");
            pres.Save(currentDir + "Table.pptx", Export.SaveFormat.Pptx);
        }
    }
}
