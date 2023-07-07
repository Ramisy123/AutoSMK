using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;


namespace AutoСМК
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        // Определение переменной oWord
        Word._Application oWord = new Word.Application();

        private void buttonDocument(object sender, EventArgs e)
        {
            // Считывает шаблон и сохраняет измененный в новом
            _Document oDoc = GetDoc(@"E:\Инст\3 курс\ТП\AutoСМК\ДокументMicrosoftWord.docx");
            string namestr = @"E:\Инст\3 курс\ТП\AutoСМК\Test" + name.Text + ".docx";
            oDoc.SaveAs(namestr);
            oDoc.Close();
        }

        public void SaveFile()
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = saveFileDialog1.FileName;
            MessageBox.Show("Файл сохранён");
        }
        

        private _Document GetDoc(string path)
        {
            _Document oDoc = oWord.Documents.Add(path);
            SetTemplate(oDoc);
            return oDoc;
        }
        // Замена закладки SECONDNAME на данные введенные в textBox
        private void SetTemplate(Word._Document oDoc)
        {
            oDoc.Bookmarks["allowanceDay"].Range.Text = allowanceDay.Text;
            oDoc.Bookmarks["allowanceMonth"].Range.Text = allowanceMonth.Text;
            oDoc.Bookmarks["allowanceYear"].Range.Text = allowanceYear.Text;
            oDoc.Bookmarks["cathedra1"].Range.Text = cathedra.Text;
            oDoc.Bookmarks["cathedra2"].Range.Text = cathedra.Text;
            oDoc.Bookmarks["course"].Range.Text = course.Text;
            oDoc.Bookmarks["deadlineDay"].Range.Text = deadlineDay.Text;
            oDoc.Bookmarks["deadlineMonth"].Range.Text = deadlineMonth.Text;
            oDoc.Bookmarks["deadlineYear"].Range.Text = deadlineYear.Text;
            oDoc.Bookmarks["defendDay"].Range.Text = defendDay.Text;
            oDoc.Bookmarks["defendMonth"].Range.Text = defendMonth.Text;
            oDoc.Bookmarks["defendYear"].Range.Text = defendYear.Text;
            oDoc.Bookmarks["discipline"].Range.Text = discipline.Text;
            oDoc.Bookmarks["fullName"].Range.Text = fullName.Text;
            oDoc.Bookmarks["group"].Range.Text = group.Text;
            oDoc.Bookmarks["name"].Range.Text = name.Text;
            oDoc.Bookmarks["name1"].Range.Text = name.Text;
            oDoc.Bookmarks["nameTeacher"].Range.Text = nameTeacher.Text;
            oDoc.Bookmarks["nameTeacher1"].Range.Text = nameTeacher.Text;
            oDoc.Bookmarks["statusTeacher"].Range.Text = statusTeacher.Text;
            oDoc.Bookmarks["theme1"].Range.Text = theme.Text;
            oDoc.Bookmarks["theme2"].Range.Text = theme.Text;
            oDoc.Bookmarks["year1"].Range.Text = Year.Text;
            oDoc.Bookmarks["year2"].Range.Text = Year.Text;
            // если нужно заменять другие закладки, тогда копируем верхнюю строку изменяя на нужные параметры 
        }
    }
}
