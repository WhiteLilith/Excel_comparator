using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_comparator
{
    class Human
    {
        private string Name;
        private string Surename;
        private string Patronymic;

        public Human(string name, string surename, string patronymic)
        {
            Name = name;
            Surename = surename;
            Patronymic = patronymic;
        }

        public Human()
        {

        }

        void SetName(string name)
        {
            Name = name;
        }

        void SetSurename(string surename)
        {
            Surename = surename;
        }

        void SetPatronymic(string patronymic)
        {
            Patronymic = patronymic;
        }

        public string GetFullName()
        {
            return (Surename + " " + Name + " " + Patronymic);
        }

        public string GetName()
        {
            return Name;
        }

        public string GetSurename()
        {
            return Surename;
        }

        public string GetPatronymic()
        {
            return Patronymic;
        }

        public static bool operator !=(Human firstHuman, Human secondHuman)
        {
            if ((firstHuman.Name == secondHuman.Name) && (firstHuman.Surename == secondHuman.Surename) && (firstHuman.Patronymic == secondHuman.Patronymic))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public static bool operator ==(Human firstHuman, Human secondHuman)
        {
            if ((firstHuman.Name == secondHuman.Name) && (firstHuman.Surename == secondHuman.Surename) && (firstHuman.Patronymic == secondHuman.Patronymic))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
