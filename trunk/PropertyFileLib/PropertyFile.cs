using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace NET.Commons
{
    public class PropertyFile
    {
        private string _FilePath;
        private Dictionary<string, string> Values = new Dictionary<string, string>();
        
        public string FilePath { get { return _FilePath; } set { _FilePath = value; } }


        private PropertyFile()
        {

        }

        private void ReadValues()
        {
            if (File.Exists(FilePath))
            {
                foreach (var row in File.ReadAllLines(FilePath))
                {
                    string[] rowData = row.Split('=');
                    Values[rowData[0]] = rowData[1];
                }
            }
            else
                throw new PropertyFileException(FilePath+" is invalid");
        }

        private void WriteValues()
        {
            StreamWriter sw = new StreamWriter(FilePath);
            try
            {
                foreach (var item in Values)
                {
                    sw.WriteLine(item.Key + "=" + item.Value);
                }
            }
            finally
            {
                sw.Close();
            }
        }

        public PropertyFile(string FilePath)
        {
            this.FilePath = FilePath;
        }

        public string Get(string PropName)
        {
            try
            {
                ReadValues();
                {
                    return Values[PropName];
                }
            }
            catch (Exception ex)
            {
                throw new PropertyFileException("Error getting property value", ex);
            }
        }

        public void Put(string PropName, string Value)
        {
            Values[PropName] = Value;
            WriteValues();
        }

        public bool Exists()
        {
            return File.Exists(FilePath);
        }

    }
}
