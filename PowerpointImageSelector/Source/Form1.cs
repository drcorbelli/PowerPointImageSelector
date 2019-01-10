//  Author:                     David Corbelli
//  Practical Coding Test:      Powerpoint generation and image search using Windows Forms 
//  Date:                       1/9/19
//

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using Syncfusion.Presentation;
using System.IO;

namespace PowerpointImageSelector
{
    public partial class Form1 : Form
    {
        UserData user;
        public Form1()
        {

            InitializeComponent();
            user = new UserData();
            pbImage1.SizeMode = PictureBoxSizeMode.Zoom;
            pbImage2.SizeMode = PictureBoxSizeMode.Zoom;
            pbImage3.SizeMode = PictureBoxSizeMode.Zoom;
            pbImage4.SizeMode = PictureBoxSizeMode.Zoom;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pbImage1_Click(object sender, EventArgs e)
        {
            if (!user.Images.Contains(pbImage1.Image) && pbImage1.Image != null)
            {
                user.Images.Add(pbImage1.Image);
                pbImage1.BorderStyle = BorderStyle.FixedSingle;
            }
            else
            {
                user.Images.Remove(pbImage1.Image);
                pbImage1.BorderStyle = BorderStyle.None;
            }
        }

        private void pbImage2_Click(object sender, EventArgs e)
        {
            if (!user.Images.Contains(pbImage2.Image) && pbImage2.Image != null)
            {
                user.Images.Add(pbImage2.Image);
                pbImage2.BorderStyle = BorderStyle.FixedSingle;
            }
            else
            {
                user.Images.Remove(pbImage2.Image);
                pbImage2.BorderStyle = BorderStyle.None;
            }
        }

        private void pbImage3_Click(object sender, EventArgs e)
        {
            if (!user.Images.Contains(pbImage3.Image) && pbImage3.Image != null)
            {
                user.Images.Add(pbImage3.Image);
                pbImage3.BorderStyle = BorderStyle.FixedSingle;
            }
            else
            {
                user.Images.Remove(pbImage3.Image);
                pbImage3.BorderStyle = BorderStyle.None;
            }
        }

        private void pbImage4_Click(object sender, EventArgs e)
        {
            if (!user.Images.Contains(pbImage4.Image) && pbImage4.Image != null)
            {
                user.Images.Add(pbImage4.Image);
                pbImage4.BorderStyle = BorderStyle.FixedSingle;
            }
            else
            {
                user.Images.Remove(pbImage4.Image);
                pbImage4.BorderStyle = BorderStyle.None;
            }
        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            user.Images.Clear();
            user.Title = txtTitle.Text;
            user.TextField = txtTextField.Text;

            string searchString = user.Title;
            foreach (string word in user.Keywords)
            {
                searchString = searchString + " " + word;
            }


            //*******************************************************************************************************************//
            // Reference Title: C# windows forms, load first google image in app itself
            // Author: Joe Shaw
            // Date: Nov 2 '16
            // Source: https://stackoverflow.com/questions/40370423/c-sharp-windows-forms-load-first-google-image-in-app-itself
            //*******************************************************************************************************************//

            // Referenced for the purpose of adapting code to bring in first image search term returns from Google 
            // into a list of four based on user input title and list of keywords

            string templateUrl = @"https://www.google.co.uk/search?q={0}&tbm=isch&site=imghp";

            if (string.IsNullOrEmpty(user.Title))
            {
                MessageBox.Show("Please supply a search term"); return;
            }
            else
            {
                using (WebClient wc = new WebClient())
                {
                    string result = wc.DownloadString(String.Format(templateUrl, new object[] { searchString }));

                    if (result.Contains("images_table"))
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(result);

                        //lets create a linq query to find all the img's stored in that images_table class.

                        var imgList = from tables in doc.DocumentNode.Descendants("table")
                                      from img in tables.Descendants("img")
                                      where tables.Attributes["class"] != null && tables.Attributes["class"].Value == "images_table"
                                      && img.Attributes["src"] != null && img.Attributes["src"].Value.Contains("images?")
                                      select img;

                        byte[] downloadedData;
                        int imgCt = 1;

                        //Adapted function to read images individually from query and correspond to forms picturebox 
                        foreach (var img in imgList)
                        {
                            if (imgCt >= 5)
                            {
                                break;
                            }
                            if (img.Attributes["src"].Value != null)
                            {
                                downloadedData = wc.DownloadData(img.Attributes["src"].Value);
                                if (downloadedData != null)
                                {
                                    System.IO.MemoryStream ms = new System.IO.MemoryStream(downloadedData, 0, downloadedData.Length);

                                    ms.Write(downloadedData, 0, downloadedData.Length);

                                    //assign memory stream to each picturebox
                                    switch (imgCt)
                                    {
                                        case 1:
                                            pbImage1.Image = Image.FromStream(ms);
                                            break;
                                        case 2:
                                            pbImage2.Image = Image.FromStream(ms);
                                            break;
                                        case 3:
                                            pbImage3.Image = Image.FromStream(ms);
                                            break;
                                        case 4:
                                            pbImage4.Image = Image.FromStream(ms);
                                            break;
                                    }

                                    imgCt++;
                                }
                            }
                        }

                    }

                }

            }
        }

        //Create slide using SyncFusion framework
        private void btnCreateSlide_Click(object sender, EventArgs e)
        {
            user.Title = txtTitle.Text;
            user.TextField = txtTextField.Text;

            IPresentation powerpointDoc = Presentation.Create();
            ISlide slide = powerpointDoc.Slides.Add(SlideLayoutType.Blank);

            IShape title = slide.AddTextBox(50, 40, 500, 100);
            title.TextBody.AddParagraph(user.Title);
            title.TextBody.Paragraphs[0].Font.FontSize = 24;


            IShape textfield = slide.AddTextBox(100, 110, 800, 300);
            textfield.TextBody.AddParagraph(user.TextField);

            for (int imgCt = 1; imgCt <= user.Images.Count; imgCt++)
            {
                var ms = new MemoryStream();
                user.Images[imgCt-1].Save(ms, ImageFormat.Png);

                IPicture picture = slide.Pictures.AddPicture(ms, ReturnImgLength(imgCt), 350, user.Images[imgCt-1].Width * 1.25, user.Images[imgCt-1].Height * 1.25);
            }

            powerpointDoc.Save("GeneratedPowerpoint.pptx");
            powerpointDoc.Close();

            ResetSelection();
        }

        //Clears all user selected information
        private void ResetSelection()
        {
            user.Images.Clear();
            user.Keywords.Clear();
            pbImage1.BorderStyle = BorderStyle.None;
            pbImage2.BorderStyle = BorderStyle.None;
            pbImage3.BorderStyle = BorderStyle.None;
            pbImage4.BorderStyle = BorderStyle.None;
            txtTextField.Font = new Font(txtTextField.Font, FontStyle.Regular);
        }

        //Space output of images based on number selected  
        private double ReturnImgLength(int imgCt)
        {
            double xCount = 0;
            for (int i = 0; i < user.Images.Count; i++)
            {
                xCount += (user.Images[i].Width*1.25);
            }
            double xspace = (975 - xCount) / (user.Images.Count+1);

            double space = 0;
            int rCount = 1;
            while (rCount <= imgCt)
            {
                if (rCount == 1)
                {
                    space += xspace;
                    rCount++;
                }
                else
                {
                    space += user.Images[rCount-2].Width*1.25 + xspace;
                    rCount++;
                }
            }
            return space;
        }

        //Bold selected text and add selected words to list of search terms
        private void btnBold_Click(Object sender, EventArgs e)
        {
            if (!txtTextField.SelectionFont.Bold && txtTextField.SelectionLength > 0) { 
                txtTextField.SelectionFont = new Font(txtTextField.Font, FontStyle.Bold | FontStyle.Regular);
                if (!user.Keywords.Contains(txtTextField.SelectedText))
                {
                    user.Keywords.Add(txtTextField.SelectedText);
                }
            }
            else
            {
                txtTextField.SelectionFont = new Font(txtTextField.Font, FontStyle.Regular);
                if (user.Keywords.Contains(txtTextField.SelectedText))
                {
                    user.Keywords.Remove(txtTextField.SelectedText);
                }
            }
        }

    }
}
