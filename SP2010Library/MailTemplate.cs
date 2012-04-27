using System;

namespace SP2010Library
{
    public class MailTemplate
    {
        public String Header;
        public String Body;
        public String Footer;
    }

    [Serializable]
    public class EmailDataPart
    {
        public string HeaderData;
        public string Value;
    }
}