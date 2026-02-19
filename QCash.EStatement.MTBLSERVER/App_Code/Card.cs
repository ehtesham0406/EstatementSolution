
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class Card
    {
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }
        private string _PAN;

        public string PAN
        {
            get { return _PAN; }
            set { _PAN = value; }
        }
        private string _MBR;

        public string MBR
        {
            get { return _MBR; }
            set { _MBR = value; }
        }
        private string _CLIENTNAME;

        public string CLIENTNAME
        {
            get { return _CLIENTNAME; }
            set { _CLIENTNAME = value; }
        }
    }
    public class CardList : List<Card> { }
}
