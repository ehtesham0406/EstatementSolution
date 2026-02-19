using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StatementGenerator
{
    public class Operation
    {
        private string _STATEMENTNO;

        public string STATEMENTNO
        {
            get { return _STATEMENTNO; }
            set { _STATEMENTNO = value; }
        }
        private string _OpID;

        public string OpID
        {
            get { return _OpID; }
            set { _OpID = value; }
        }
        private string _OpDate;

        public string OpDate
        {
            get { return _OpDate; }
            set { _OpDate = value; }
        }
        private string _TD;

        public string TD
        {
            get { return _TD; }
            set { _TD = value; }
        }
        private string _Amount;

        public string Amount
        {
            get { return _Amount; }
            set { _Amount = value; }
        }
        private string _ACURCode;

        public string ACURCode
        {
            get { return _ACURCode; }
            set { _ACURCode = value; }
        }
        private string _ACURName;

        public string ACURName
        {
            get { return _ACURName; }
            set { _ACURName = value; }
        }
        private string _D;

        public string D
        {
            get { return _D; }
            set { _D = value; }
        }
        private string _DE;

        public string DE
        {
            get { return _DE; }
            set { _DE = value; }
        }
        private string _P;

        public string P
        {
            get { return _P; }
            set { _P = value; }
        }
        private string _OA;

        public string OA
        {
            get { return _OA; }
            set { _OA = value; }
        }
        private string _OCCode;

        public string OCCode
        {
            get { return _OCCode; }
            set { _OCCode = value; }
        }
        private string _OCName;

        public string OCName
        {
            get { return _OCName; }
            set { _OCName = value; }
        }
        private string _TL;

        public string TL
        {
            get { return _TL; }
            set { _TL = value; }
        }
        private string _TERMN;

        public string TERMN
        {
            get { return _TERMN; }
            set { _TERMN = value; }
        }
        private string _CF;

        public string CF
        {
            get { return _CF; }
            set { _CF = value; }
        }
        private string _S;

        public string S
        {
            get { return _S; }
            set { _S = value; }
        }
        private string _MN;

        public string MN
        {
            get { return _MN; }
            set { _MN = value; }
        }
        private string _DOCNO;

        public string DOCNO
        {
            get { return _DOCNO; }
            set { _DOCNO = value; }
        }
        private string _NO;

        public string NO
        {
            get { return _NO; }
            set { _NO = value; }
        }
        private string _ACCOUNT;

        public string ACCOUNT
        {
            get { return _ACCOUNT; }
            set { _ACCOUNT = value; }
        }
        private string _ACC;

        public string ACC
        {
            get { return _ACC; }
            set { _ACC = value; }
        }
        private string _FR;

        public string FR
        {
            get { return _FR; }
            set { _FR = value; }
        }
        private string _APPROVAL;

        public string APPROVAL
        {
            get { return _APPROVAL; }
            set { _APPROVAL = value; }
        }
        private string _AMOUNTSIGN;

        public string AMOUNTSIGN
        {
            get { return _AMOUNTSIGN; }
            set { _AMOUNTSIGN = value; }
        }

        private string _SERIALNO;

        public string SERIALNO
        {
            get { return _SERIALNO; }
            set { _SERIALNO = value; }
        }
    }
    public class OperationList : List<Operation>
    { }
}
