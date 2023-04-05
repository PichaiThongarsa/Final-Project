using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Filleenprojact
{
    internal class Movie : Main
    {
        private string duration;

        public Movie(string name, string author, string category, string duration, string day) : base(name, author, category, day)
        {
            this.duration = duration;
        }
        public string getDuration()
        {
            return duration;
        }
    }
}
