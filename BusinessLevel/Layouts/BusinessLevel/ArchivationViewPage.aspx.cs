using System;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Client;

using System.Web.UI;       
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Microsoft.Office.DocumentManagement.DocumentSets;
using Microsoft.Office.Word.Server.Conversions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace _1Plus1Archivation.Layouts.BusinessLevel
{
    public partial class ArchivationViewPage : LayoutsPageBase
    {
        public string SelectedItems;
        public string ActionId;

        protected void Page_Load(object sender, EventArgs e)
        {
            // считывание заголовка запроса 
            ActionId = Request["Action"];      //   кнопка активации
            SelectedItems = Request["Ids"];    //   набор данных

            // 0 - если запрос отправлен кнопкой OnSelected
            // 1 - если запрос отправлен кнопкой OnDated
            if ((ActionId == "0") || (SelectedItems != null))
                // включение визуализации контента модального окна для двух случаев активации
            {
                h2Element.Visible = false;
                ArchiveItemsDate.Visible = false;
                ItemIDs.Text = SelectedItems;
            }
            else if (ActionId == "1")
            {
                h2Element.Visible = true;
                ArchiveItemsDate.Visible = true;
            }
        }

        // обработка контрола календаря
        protected void ArchiveItemsDate_SelectionChanged(object sender, EventArgs e)
        {
            // адрес сайта
            String sharePointSite = "http://vizatech.westeurope.cloudapp.azure.com/sites/team/";

            // получаем конткст списка
            SPSite oSite = new SPSite(sharePointSite);
            SPWeb oWeb = oSite.OpenWeb();
            SPList list = oWeb.Lists["Doc2"];

            // считываем дату с контрола
            DateTime endDate = ArchiveItemsDate.SelectedDate;

            // получаем коллекцию АйДи элементов, 
            // для которых дата завершения договора - старше выбранной пользователем даты 
            string IDs = "";
            foreach (SPListItem item in list.Items)
            {
                if (item.Fields.ContainsField("Дата завершення дії договору"))
                {
                    try
                    {
                        if ((DateTime.Compare((DateTime)item["Дата завершення дії договору"], endDate) <= 0))
                        {
                            IDs += item.ID + " ";
                        }
                    }
                    catch {}
                }
            }

            // показываем пользователю - какие элементы будут архивированы
            ItemIDs.Text = IDs;
        }

        // Активация архивирования 
        protected void ArchiveItems_Click(object sender, EventArgs e)
        {
            // адрес сайта
            String sharePointSite = "http://vizatech.westeurope.cloudapp.azure.com/sites/team/";

            // получаем контекст списка
            SPSite oSite = new SPSite(sharePointSite);
            SPWeb oWeb = oSite.OpenWeb();
            SPList list = oWeb.Lists["Doc2"];

            // получает набор АйДи для архивации
            char[] delimiterChars = { ' ', ',', '.', ':', '|' };
            string[] SelectedItemIDs = ItemIDs.Text.Split(delimiterChars);

            //  Создаем Отчет об орхивировании для каждого отобранного Айтема   
            foreach (string ID in SelectedItemIDs)
            {
                int SelectedItemIDint = 0;
                try { SelectedItemIDint = Int32.Parse(ID); }
                catch { }
                if ((SelectedItemIDint > 0))
                {
                    // полчаем контекст Айтема
                    SPListItem SelectedItem = list.Items.GetItemById(SelectedItemIDint);
                    // передаем данные методу создания отчета
                    CreateDocument(SelectedItem);
                }
            }

            oSite = new SPSite(sharePointSite);
            oWeb = oSite.OpenWeb();
            oWeb.AllowUnsafeUpdates = true;

            // получаем контекст Службы конвертации 
            ConversionJob myJob = new ConversionJob("Word Automqtion Service");

            // задаем параметры конвертации
            myJob.Settings.OutputFormat = SaveFormat.PDF;
            myJob.Settings.OutputSaveBehavior = SaveBehavior.AlwaysOverwrite;
            myJob.UserToken = oWeb.CurrentUser.UserToken;

            // задаем от куда брать и куда переносить файлы
            SPList inputLibrary = oWeb.Lists["QueueLib"];
            SPList outputLibrary = oWeb.Lists["ArchiveLib"];
            myJob.AddLibrary(inputLibrary, outputLibrary);
            myJob.Start();

            //ClearLibrary(inputLibrary);  
            // возвращаем управление вызвавшему сайту 
            Response.Redirect(sharePointSite);
        }

        // Удаление папки, если в ней нет элементов
        protected void ClearLibrary(SPList Library)
        {
            if (Library.Folders.Count > 0)
                foreach (SPListItem ClearFolder in Library.Folders)
                {
                    ClearFolder.Delete();
                    if (Library.Folders.Count == 0) break;
                }
        }

        // создание отчета
        protected void CreateDocument(SPListItem SelectedItem)
        {
            // адрес сайта
            String sharePointSite = "http://vizatech.westeurope.cloudapp.azure.com/sites/team";

            // получение контекста сайта
            SPSite oSite = new SPSite(sharePointSite);
            SPWeb oWeb = oSite.OpenWeb();

            // получение имени текущего пользователя
            string UserName = oWeb.CurrentUser.Name;

            // Создаем словарь заполнения шаблона отчета об архивации
            Dictionary<string, string> AddElementsDictionary = GetItemFields(SelectedItem, UserName);
            String TemplateLibraryName = "TemplateLib";
            String TemplateFileName = "DocTemplate.dotx";
            String ArchiveLibraryName = "QueueLib";
            String ArchiveFolderName = AddElementsDictionary.Values.ElementAt(0) + "_" + AddElementsDictionary.Values.ElementAt(4);
            String ArchiveFileName = ArchiveFolderName + ".dotx";
      
            oWeb.AllowUnsafeUpdates = true;
            // получаем контекст файла
            SPFile TemplateFile = (SPFile)oWeb.GetFileOrFolderObject(sharePointSite + "/Lists/" + TemplateLibraryName + "/" + TemplateFileName);
            // считываем файл в стрим
            Stream TemplateFileStream = new MemoryStream(TemplateFile.OpenBinary());
            // Заполняем поля шаблона данными 
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(TemplateFileStream, true))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                int CurrentBookMarkId = 0;
                // перебираем метки шаблона
                foreach (BookmarkStart bookmarkStart in mainPart.RootElement.Descendants<BookmarkStart>())
                {
                    if (CurrentBookMarkId < AddElementsDictionary.Count)
                    {
                        if (AddElementsDictionary.ElementAt(CurrentBookMarkId).Key == bookmarkStart.Name)
                        {
                            // задаем свойства вносимого текста
                            RunProperties rPr = new RunProperties(new RunFonts() { Ascii = "Arial" }, new Bold(), new Color() { Val = "green" });
                            // вносим данные в шаблон
                            Run InsertToBookmarkOperation = new Run(new Text(AddElementsDictionary.ElementAt(CurrentBookMarkId).Value));
                            InsertToBookmarkOperation.PrependChild<RunProperties>(rPr);
                            bookmarkStart.Parent.InsertAfter(InsertToBookmarkOperation, bookmarkStart);
                            CurrentBookMarkId++;
                        }
                    }
                }
                // Сохраняем изменения
                mainPart.Document.Save();
            }

            // получаем контекст папки архива и создаем папку для архивированных документов 
            SPFolder TrueArchiveLibrary = oWeb.Folders["Lists"].SubFolders["ArchiveLib"];
            TrueArchiveLibrary.SubFolders.Add(ArchiveFolderName);

            // получаем контекст папки очереди на архивирование и создаем папку 
            SPFolder ArchiveLibrary = oWeb.Folders["Lists"].SubFolders[ArchiveLibraryName];
            ArchiveLibrary.SubFolders.Add(ArchiveFolderName);

            // переносим заполненный отчет в папку очереди
            ArchiveLibrary.SubFolders[ArchiveFolderName].Files.Add(ArchiveFileName, TemplateFileStream, true);

            // инициируем архивирование файлом доксета
            // получаем контекст доксета
            DocumentSet ItemDocSet = DocumentSet.GetDocumentSet(SelectedItem.Folder);

            // проверяем - может ли быть обработан данный файл службой конвертации
            char[] chars = { '.' };
            foreach (SPFile file in ItemDocSet.Folder.Files)
            {
                Stream DocSetFileStream = new MemoryStream(file.OpenBinary());
                if (new string[] { "docx", "docm", "dotx", "dotm", "doc", "dot", "rtf", "htm", "html", "mht", "mhtml", "xml" }.Contains(file.Name.Split(chars).ElementAt(1)))
                    // если файл относится к группе конвертируемых, то он сатвится в очередь
                    ArchiveLibrary.SubFolders[ArchiveFolderName].Files.Add(file.Name, DocSetFileStream, true);
                    // иначе - сразу перенсится в папку обработанных документов без обработки
                else { TrueArchiveLibrary.SubFolders[ArchiveFolderName].Files.Add(file.Name, DocSetFileStream, true); }
            }
                                                                                             
            /*
                SPFolder AttachmentFolder = (SPFolder)oWeb.GetFileOrFolderObject(
                        oSite.Url + "/Lists/Doc2/Attachments/" + SelectedItem["ID"]);

                foreach (SPFile AttachmentFile in AttachmentFolder.Files)
                {
                    AttachmentFile.CopyTo(ArchiveLibrary.SubFolders[ArchiveFolderName].Url + '/' + AttachmentFile.Name);
                }       
            */
        }

        // формируем словарь заполнения шаблона отчета
        private static Dictionary<string, string> GetItemFields(SPListItem a, string UserName)
        {
            return new Dictionary<string, string> {
                { "DocID", (string)a["ІД проекту"]},
                { "DocType", (string)a["Вид договору"]},
                { "DocNumber", (string)a["Номер договору"]},
                { "DocRegDate", ((DateTime)a["Дата реєстрації"]).ToShortDateString()},
                { "DocDep", (string)a["Департамент"]},
                { "DocSub", (string)a["Предмет договору"]},
                { "OurSide", (string)a["Наша сторона"]},
                { "Contragent", (string)a["Контрагент"]},
                { "Manager", (string)a["Менеджер договору"]},
                { "DocSignDate", ((DateTime)a["Дата заключення"]).ToShortDateString()},
                { "Status", (string)a["Статус ЖЦ"]},
                { "Lower1", (string)a["Юрист"]},
                { "BookKeeper1", (string)a["Бухгалтер"]},
                { "Finance1", (string)a["Фінансист"]},
                { "FinDirector1", (string)a["Фін.дир./Зам.фін.дир"]},
                { "Lower2", (string)a["Юрист2"]},
                { "BookKeeper2", (string)a["Бухгалтер2"]},
                { "Finance2", (string)a["Фінансист2"]},
                { "FinDirector2", (string)a["Фін.дир./Зам.фін.дир2"]},
                { "CardCreatedDate", ((DateTime)a["Created"]).ToShortDateString()},
                { "Author", (string)a["Author"]},
                { "Project", (string)a["Проект"]},
                { "ArchBy", UserName },
                { "ArchDate", DateTime.Today.ToShortDateString()}
            };
        }
    }
}
