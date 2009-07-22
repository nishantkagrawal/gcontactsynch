using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NET.Commons
{
    public class PropertyFileException : Exception
    {
        public PropertyFileException() : base() {}
        public PropertyFileException (string Msg) : base(Msg) {}
        public PropertyFileException(string Msg, Exception ex) : base(Msg,ex) { }
    }
}
