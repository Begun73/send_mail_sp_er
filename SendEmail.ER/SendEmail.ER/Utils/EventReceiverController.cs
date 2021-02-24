using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace SendEmail.ER.Utils
{
    /// <summary>
    /// Позволяет добавлять и удалять EventReceiver к спискам SharePoint
    /// </summary>
    public static class EventReceiverController
    {
        /// <summary>
        /// Добавить списку <paramref name="list"/> EventReceiver с именем <paramref name="name"/>.
        /// </summary>
        /// <param name="name">User friendly name for EventReceiver</param>
        /// <param name="list">SPList, к которому добавляется EventReceiver </param>
        /// <param name="typeClassReceiver">Тип (класс), в котором реализован метод-обработчик соответствующего EventReceiver'а</param>
        /// <param name="typeReceiver">Тип EventReceiver'а. Используйте только типы EventReceiver'ов, которые используются для СПИСКОВ (SPList)</param>
        /// <param name="synchronizationValue" >Синхронное или Асинхронное</param>
        /// <param name="sequenceNumber">Число, представляющее место данного события в последовательности событий</param>
        /// <exception cref="ArgumentNullException">Возникает если следующие параметры равны null: <paramref name="name"/>, <paramref name="list"/>, <paramref name="typeClassReceiver"/>.
        /// </exception>
        /// <exception cref="ArgumentException">Возникает если у списка уже есть Event ресивер типа <paramref name="typeReceiver"/> реализованный с помощью класса <paramref name="typeClassReceiver"/>
        /// </exception>
        public static void AddListEventReceiver(string name, SPList list, Type typeClassReceiver, SPEventReceiverType typeReceiver, SPEventReceiverSynchronization synchronizationValue, int sequenceNumber)
        {
            #region Проверка входящих параметров на null

            if (String.IsNullOrEmpty(name))
            {
                throw new ArgumentNullException("name");
            }

            if (list == null)
            {
                throw new ArgumentNullException("list");
            }

            if (typeClassReceiver == null)
            {
                throw new ArgumentNullException("typeClassReceiver");
            }

            #endregion

            var eventReceiverAssembly = typeClassReceiver.Assembly.FullName;
            var eventReceiverClass = typeClassReceiver.FullName;

            #region Проверяем есть ли уже в списке такой Event Reciever.

            for (var i = 0; i < list.EventReceivers.Count; i++)
            {
                var eventReceiverDefinition = list.EventReceivers[i];
                if (String.Equals(eventReceiverDefinition.Name, name))
                {
                    throw new ArgumentException("Event Receiver с таким именем уже существует.", "name");
                }

                if (eventReceiverDefinition.Assembly == eventReceiverAssembly && eventReceiverDefinition.Class == eventReceiverClass && eventReceiverDefinition.Type == typeReceiver)
                {
                    throw new ArgumentException(
                        String.Format("Такой Event Receiver уже существует. eventReceiverClass = {0} eventReceiverAssembly = {1}, typeReceiver = {2}", eventReceiverClass, eventReceiverAssembly, typeReceiver));
                }
            }

            #endregion

            // Создаём новый EventReceiver
            SPEventReceiverDefinition newEventReceiverDefinition = list.EventReceivers.Add();
            newEventReceiverDefinition.Type = typeReceiver;
            newEventReceiverDefinition.Assembly = typeClassReceiver.Assembly.FullName;
            newEventReceiverDefinition.Class = typeClassReceiver.FullName;
            // Задаём правильное имя EventReceiver'у 
            newEventReceiverDefinition.Name = name;
            // Задаём тип синхронизации
            newEventReceiverDefinition.Synchronization = synchronizationValue;

            newEventReceiverDefinition.SequenceNumber = sequenceNumber;

            newEventReceiverDefinition.Update();
        }

        /// <summary>
        /// Удалить EventReceiver типа <paramref name="typeReceiver"/> у списка <paramref name="list"/> , который реализованн с помощью класса <paramref name="typeClassReceiver"/>.
        /// Если такого EventReceiver'а не найдено, то ошибка генерироваться НЕ будет!
        /// </summary>
        /// <param name="list"></param>
        /// <param name="typeClassReceiver"></param>
        /// <param name="typeReceiver"></param>
        /// <exception cref="ArgumentNullException">Возникает если следующие параметры равны null: <paramref name="list"/>, <paramref name="typeClassReceiver"/>.
        /// </exception>
        public static void DeleteListEventReceiver(SPList list, Type typeClassReceiver, SPEventReceiverType typeReceiver)
        {
            #region Проверка входящих параметров на null

            if (list == null)
            {
                throw new ArgumentNullException("list");
            }

            if (typeClassReceiver == null)
            {
                throw new ArgumentNullException("typeClassReceiver");
            }

            #endregion

            var count = list.EventReceivers.Count;

            for (var i = 0; i < count; i++)
            {
                var eventReceiver = list.EventReceivers[i];
                if (eventReceiver.Assembly == typeClassReceiver.Assembly.FullName &&
                    eventReceiver.Class == typeClassReceiver.FullName &&
                    eventReceiver.Type == typeReceiver)
                {
                    eventReceiver.Delete();
                    i = i - 1;
                    count = count - 1;
                }
            }
        }

        /// <summary>
        /// Обеспечить наличие на списке <paramref name="list"/> EventReceiver'а <paramref name="name"/>.
        /// Этот метод позволяет программно повесить обработчик события на список.
        /// Если на момент вызова данного метода, такой обработчик на списке уже висел, то он будет повешен заново.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="list"></param>
        /// <param name="typeClassReceiver"></param>
        /// <param name="typeReceiver"></param>
        /// <param name="synchronizationValue"></param>
        /// <param name="sequenceNumber"></param>
        public static void ProvisionListEventReceiver(string name, SPList list, Type typeClassReceiver, SPEventReceiverType typeReceiver, SPEventReceiverSynchronization synchronizationValue, int sequenceNumber)
        {
            DeleteListEventReceiver(list, typeClassReceiver, typeReceiver);
            AddListEventReceiver(name, list, typeClassReceiver, typeReceiver, synchronizationValue, sequenceNumber);
        }

    }
}
