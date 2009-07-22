using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Extensions;

namespace GContactsSync
{
    public class EmailComparer : IEqualityComparer<EMail>
    {
        #region IEqualityComparer<EMail> Members

        public bool Equals(EMail x, EMail y)
        {
            return x.Address==y.Address;
        }

        public int GetHashCode(EMail obj)
        {
            return obj.GetHashCode();
        }

        #endregion
    }
}
