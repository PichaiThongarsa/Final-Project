using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Filleenprojact
{
    internal class Main
    {
        private string name;
        private string author;
        private string category;
        private string day;

        public Main(string name, string author, string category, string day)
        {
            this.name = name;
            this.author = author;
            this.category = category;
            this.day = day;
        }
        public string getName()
        {
            return name;
        }
        public string getAuthor()
        {
            return author;
        }
        public string getCategory()
        {
            return category;
        }
        public string getDay()
        {
            return day;
        }
    }
}