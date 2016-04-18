using Convertion;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //ConvertToPDF();
        }

        private void ConvertToImages()
        {

            string path = this.txtsourceFile.Text;
            ConvertFromPPTX pptx = new ConvertFromPPTX();
            bool flag = pptx.ConvertToImages(path, this.txtTargetFolder.Text);
            if (flag)
                MessageBox.Show("转换成功");
            else
            {
                MessageBox.Show("转换失败");
            }
            
        }

        private void ConvertToPDF()
        {
            string path = this.txtsourceFile.Text;
            ConvertFromPPTX pptx = new ConvertFromPPTX();
            bool flag = pptx.ConvertToPdf(path, this.txtTargetFolder.Text);
            if (flag)
                MessageBox.Show("转换成功");
            else
            {
                MessageBox.Show("转换失败");
            }
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            string filePath = String.Empty; string fileName = String.Empty; string fileExtension = String.Empty;
            string[] strExtension = new string[] { ".docx", ".pptx"}; 
            //openFileDialog1.Filter = "All files (*.*)|*.*";
            openFileDialog1.Title = "Choose Office Doc File";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //文件路径 和文件名字  
                filePath = openFileDialog1.FileName;
                //fileName = Path.GetFileName(filePath);
                // 获取文件后缀  
                fileExtension = Path.GetExtension(filePath);
                //设置一个text控件，显示更新txtShow的显示内容  
                txtsourceFile.Text = filePath;

                if (!strExtension.Contains(fileExtension))//验证读取文件的格式，设置为只能读取以下几种格式的图片  
                {
                    MessageBox.Show("仅打开.docx, .pptx格式的文档");
                }
            }
        }

        private void btnChooseFolder_Click(object sender, EventArgs e)
        {
            string filePath = String.Empty; string fileName = String.Empty; string fileExtension = String.Empty;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                //文件路径 和文件名字  
                filePath = folderBrowserDialog1.SelectedPath;
                
                //设置一个text控件，显示更新txtShow的显示内容  
                txtTargetFolder.Text = filePath;

            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text == "To Image")
                ConvertToImages();
            else if (this.comboBox1.Text == "To PDF")
                ConvertToPDF();
        }

    }
}
