using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace _1PLUS1ArchivationEventReceiver.ArchivationEventReceiver
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class ArchivationEventReceiver : SPItemEventReceiver
    {
        /// <summary>
        /// An item is being added.
        /// </summary>
        public override void ItemAdding(SPItemEventProperties properties)
        {
            // получаем контекст добавленного Айтема
            SPListItem item = properties.ListItem;

            // удаляем родительский файл в папке очереди



            // если папка очереди пуста, то удаляем папку и проверяем создан ли в отчет об архивации в папке архивации


                 // если отчет существует - удаляем запись в списке главной библиотеки


            base.ItemAdding(properties);
        }
    }
}