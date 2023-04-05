using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Filleenprojact
{
    internal class Book : Main
    {
        private string page;

        public Book(string name, string author, string category,  string page , string day) : base(name,author,category,day)
        {
            this.page = page;
        }
        public string getPage()
        {
            return page;
        }
    }
}
