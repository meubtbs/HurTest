using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace HurTest
{
    public partial class ThisDocument
    {
        private string docDataSetID = null;
        private HurTestDataSet docDataSet = new HurTestDataSet();
 
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            this.ContentControlOnEnter += ThisDocument_ContentControlOnEnter;
            this.BeforeSave += ThisDocument_BeforeSave;
            try
            {
                Office.DocumentProperties docProps =
                    this.CustomDocumentProperties as Office.DocumentProperties;
                foreach (Office.DocumentProperty docprop in docProps)
                {
                    if(docprop.Name == "HurTestDataSet")
                    {
                        docDataSetID = (string)docprop.Value;
                        if (docDataSetID != null)
                        {
                            Office.CustomXMLPart docDataSetPart =
                                this.CustomXMLParts.SelectByID(docDataSetID);
                            System.IO.StringReader srdr = new System.IO.StringReader(docDataSetPart.XML);
                            docDataSet.ReadXml(srdr);
                        }
                    }
                }          
            }
            catch
            {

            }
        }

        void ThisDocument_ContentControlOnEnter(Word.ContentControl ContentControl)
        {
            string ctrlID = ContentControl.ID;
        }

        private void ThisDocument_BeforeSave(object sender, Microsoft.Office.Tools.Word.SaveEventArgs e)
        {
            Office.CustomXMLPart docDataSetPart = null;
            // Belgedeki içerik kontrol kutularıyla ilgili bilgileri içeren
            // veri setini belgeye ait XML alanları arasında sakla.
            if (docDataSetID != null)
            {
                docDataSetPart = this.CustomXMLParts.SelectByID(docDataSetID);
                if(docDataSetPart != null) docDataSetPart.Delete();
            }
            string xmlDataSet = docDataSet.GetXml();
            docDataSetPart = this.CustomXMLParts.Add(xmlDataSet);
            Office.DocumentProperties docProps =
                this.CustomDocumentProperties as Office.DocumentProperties;
            if (docDataSetID != null) 
            {
                docDataSetID = docDataSetPart.Id;
                foreach (Office.DocumentProperty docprop in docProps)
                {
                    if (docprop.Name == "HurTestDataSet")
                    {
                        docprop.Value = docDataSetID;
                        break;
                    }
                }
            }
            else 
            {
               docDataSetID = docDataSetPart.Id;
               docProps.Add("HurTestDataSet", false, Office.MsoDocProperties.msoPropertyTypeString,
               docDataSetID, missing);
            }          
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
           
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion

        /// <summary>
        /// Bu metot belgeye içine soru metinler için zengin metin kutucukları konacak
        /// olan bir tablo içeren bir zengin metin kutusu ekler.
        /// </summary>
        public void SoruGrubuEkle()
        {
            // Soru grubu kutusu belge sonuna eklenecek
            object end = this.Content.End-1;
            Word.ContentControl ctrlGroup = this.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                this.Range(ref end, ref end));
            HurTestDataSet.QuestionGroupsRow qgroupRow =
                docDataSet.QuestionGroups.AddQuestionGroupsRow(ctrlGroup.ID, 1, 1);
            // Sorular için bu kutu içinde bir tablo olacak.
            Word.Table tblGroup = this.Tables.Add(ctrlGroup.Range, 1, 2);
            tblGroup.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tblGroup.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth225pt;
            tblGroup.Borders.OutsideColor = Word.WdColor.wdColorBlue;
            tblGroup.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            tblGroup.Borders.InsideLineWidth = Word.WdLineWidth.wdLineWidth025pt;
            tblGroup.Borders.InsideColor = Word.WdColor.wdColorBlue;
            tblGroup.Columns[1].SetWidth(50, Word.WdRulerStyle.wdAdjustFirstColumn);
            object groupbegin = ctrlGroup.Range.Start;

            this.Range(ref groupbegin, ref groupbegin).InsertBefore("Seçilecek Soru Sayısı: ");
            this.Range(ref groupbegin, ref groupbegin).InsertParagraph();
            Word.Range par1range = ctrlGroup.Range.Paragraphs[1].Range;
            // Soru grubu kutusunda açıklamalar için bir metin kutusu olacak.
            Word.ContentControl ctrlComment = this.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText, par1range);
            ctrlComment.SetPlaceholderText(null, null, "Soru grubu açıklamalarını buraya yazın");
            ctrlComment.LockContentControl = true;
            // Soru grubu kutusunda bir de seçilecek soru sayısını belirlemek
            // için bir açılır liste kutusu olacak.
            object par2end = ctrlGroup.Range.Paragraphs[2].Range.End - 1;
            Word.ContentControl cbQuestionCount = this.ContentControls.Add(
                Word.WdContentControlType.wdContentControlComboBox,
                this.Range(ref par2end, ref par2end));
            cbQuestionCount.DropdownListEntries.Add("1","1",0);
            cbQuestionCount.SetPlaceholderText(null, null, "Sayı seçin");
            cbQuestionCount.LockContentControl = true;
            // Gruba eklenen ilk boş soru için seçim kutusu konacak.
            Word.ContentControl chkIncludeQuestion = this.ContentControls.Add(
                Word.WdContentControlType.wdContentControlCheckBox,
                tblGroup.Cell(1,1).Range);
            chkIncludeQuestion.Checked = true;
            chkIncludeQuestion.LockContentControl = true;
            // Gruba bir boş soru kutusu konacak.
            Word.ContentControl ctrlQuestion = this.ContentControls.Add(
                Word.WdContentControlType.wdContentControlRichText,
                tblGroup.Cell(1, 2).Range);
            docDataSet.Questions.AddQuestionsRow(ctrlQuestion.ID, qgroupRow, chkIncludeQuestion.ID, true, 0);
            ctrlQuestion.SetPlaceholderText(null, null, "Soru metnini buraya yazın");
            ctrlQuestion.LockContentControl = true;
        }

        /// <summary>
        /// Bu metot o an aktif olan soru grubu kutusundaki tabloya
        /// içine bir soru metni yazılacak bir zengin metin kutusu 
        /// ve o sorunun için bir seçim kutusu ekler.
        /// </summary>
        public void SoruEkle()
        {
            // Her bir form kutusunu tara ve o bir soru grubu mu diye bak.
            foreach (Word.ContentControl ctrlGroup in this.ContentControls)
            {
                HurTestDataSet.QuestionGroupsRow qgroupRow =
                    docDataSet.QuestionGroups.FindByGroupID(ctrlGroup.ID);
                if (qgroupRow == null) continue;
                // O an seçili konum hangi soru grubu kutusundaysa o kutudaki
                // tabloya bir soru kutusu ekle.
                if(Globals.ThisDocument.Application.Selection.Range.InRange(ctrlGroup.Range))
                {
                    Word.Table tblGroup = ctrlGroup.Range.Tables[1];
                    tblGroup.Rows.Add();

                    Word.ContentControl chkIncludeQuestion = this.ContentControls.Add(
                        Word.WdContentControlType.wdContentControlCheckBox,
                        tblGroup.Rows.Last.Cells[1].Range);
                    chkIncludeQuestion.Checked = true;
                    chkIncludeQuestion.LockContentControl = true;

                    Word.ContentControl ctrlQuestion = this.ContentControls.Add(
                        Word.WdContentControlType.wdContentControlRichText,
                        tblGroup.Rows.Last.Cells[2].Range);
                    docDataSet.Questions.AddQuestionsRow(ctrlQuestion.ID, qgroupRow, chkIncludeQuestion.ID, true, 0);
                    ctrlQuestion.SetPlaceholderText(null, null, "Soru metnini buraya yazın");
                    ctrlQuestion.LockContentControl = true;
                }
            }
        }

        /// <summary>
        /// Bu metot o an aktif olan soru kutusunun sonuna içine bir seçenek
        /// metni yazılacak bir zengin metin kutusu ekler.
        /// </summary>
        public void InsertChoiceField()
        {
            // Her bir form kutusunu tara ve o bir soru kutusu mu diye bak.
            foreach (Word.ContentControl ctrlQuestion in this.ContentControls)
            {
                HurTestDataSet.QuestionsRow questionRow =
                    docDataSet.Questions.FindByQuestionID(ctrlQuestion.ID);
                if (questionRow == null) continue;
                
                Word.Range rngSelected = Globals.ThisDocument.Application.Selection.Range;
                // O an seçili konum hangi soru  kutusundaysa o kutudaki
                // tabloya bir seçenek kutusu ekle.
                if (ctrlQuestion.Range.Start <= rngSelected.Start &&
                    ctrlQuestion.Range.End >= rngSelected.End)
                {
                    ctrlQuestion.Range.InsertParagraphAfter();

                    if (questionRow.ChoiceCount == 0)
                    { ctrlQuestion.Range.InsertParagraphAfter(); }

                    object posChoice = ctrlQuestion.Range.End-1;
                    Word.Range rngChoice = this.Range(ref posChoice, ref posChoice);
                    Word.ContentControl ctrlChoice = this.ContentControls.Add(
                        Word.WdContentControlType.wdContentControlRichText, rngChoice);
                    ctrlChoice.SetPlaceholderText(null, null, "Seçenek metnini buraya yazın");

                    questionRow.ChoiceCount = questionRow.ChoiceCount + 1;
                    docDataSet.Choices.AddChoicesRow(ctrlChoice.ID, questionRow, false, false);
                }
            }
        }

        /// <summary>
        /// Bu metot o an aktif olan soru kutusunun sonuna içinde bir değişken
        ///sayısal değer gözükecek olan bir zengin metin kutusu ekler.
        /// </summary>
        public void InsertNumericField()
        {
            foreach (Word.ContentControl ctrlQuestion in this.ContentControls)
            {
                if (docDataSet.Questions.FindByQuestionID(ctrlQuestion.ID) == null) continue;

                Word.Range rngSelected = Globals.ThisDocument.Application.Selection.Range;
                if (ctrlQuestion.Range.Start <= rngSelected.Start &&
                    ctrlQuestion.Range.End >= rngSelected.End)
                {
                    Word.ContentControl ctrlNumeric = this.ContentControls.Add(
                        Word.WdContentControlType.wdContentControlText, rngSelected);
                    ctrlNumeric.SetPlaceholderText(null,null,"<1|2|3>");
                    ctrlNumeric.LockContents = true;

                    docDataSet.Numerics.AddNumericsRow(ctrlNumeric.ID, 0, 0, 0);
                }
            }
        }
    }
}
