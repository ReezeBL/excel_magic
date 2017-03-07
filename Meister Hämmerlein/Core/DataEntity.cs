using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Meister_Hämmerlein.Annotations;
using CurrencyManager = Meister_Hämmerlein.Core.СurrencyManager;

namespace Meister_Hämmerlein.Core
{
    public class DataEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private DateTime date;
        private string document;
        private string analytics;
        private decimal debetRu;
        private decimal debetUsd;
        private decimal creditRu;
        private decimal creditUsd;
        private string type;

        [DisplayName("Период")]
        public DateTime Date
        {
            get { return date; }
            set {
                if (date == value)
                    return;
                date = value;
                OnPropertyChanged();
            }
        }
        [DisplayName("Документ")]
        public string Document
        {
            get { return document; }
            set
            {
                if (document == value)
                    return;
                document = value;
                OnPropertyChanged();
            }
        }
        [DisplayName("Аналитика Дт")]
        public string Analytics
        {
            get { return analytics; }
            set
            {
                if (analytics == value)
                    return;
                analytics = value;
                OnPropertyChanged();
            }
        }
        [DisplayName("Дебет, рубли")]
        public decimal DebetRu
        {
            get { return debetRu; }
            set
            {
                if (debetRu == value)
                    return;
                debetRu = value;
                OnPropertyChanged();

                var rate = CurrencyManager.GetRate(date);
                debetUsd = Math.Round(value / rate, 2);
                OnPropertyChanged(nameof(DebetUsd));
            }
        }
        [DisplayName("Дебет, USD")]
        public decimal DebetUsd
        {
            get { return debetUsd; }
            set
            {
                if (debetUsd == value)
                    return;
                debetUsd = value; 
                OnPropertyChanged();

                var rate = CurrencyManager.GetRate(date);
                debetRu = value * rate;
                OnPropertyChanged(nameof(DebetRu));
            }
        }
        [DisplayName("Кредит, рубли")]
        public decimal CreditRu
        {
            get { return creditRu; }
            set
            {
                if (creditRu == value)
                    return;
                creditRu = value;
                OnPropertyChanged();

                var rate = CurrencyManager.GetRate(date);
                creditUsd = Math.Round(value / rate, 2);
                OnPropertyChanged(nameof(CreditUsd));
            }
        }
        [DisplayName("Кредит, USD")]
        public decimal CreditUsd
        {
            get { return creditUsd; }
            set
            {
                if(creditUsd == value)
                    return;
                creditUsd = value;
                OnPropertyChanged();

                var rate = CurrencyManager.GetRate(date);
                creditRu = value * rate;
                OnPropertyChanged(nameof(CreditRu));
            }
        }

        [DisplayName("Тип")]
        public string Type
        {
            get { return type; }
            set
            {
                if (type == value)
                    return;
                type = value;
                OnPropertyChanged();
            }
        }


        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}