using System;
using System.Collections.Specialized;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using System.Reflection;
//
namespace FlexiStar.Utilities.EncryptionEngine
{
    /// <summary>
    /// Summary description for Encryption.
    /// Summary description for Encryption.
    /// //
    /// </summary>
    
    public class Encryption
    {
        // Fields
        private bool IsValidUser = false;
        protected StringCollection stringCollection;

        // Methods
        public Encryption()
        {
            this.PopulateErrorString();
            this.IsValidUser = true;
        }

        public string DecryptWord(string DecryptIt)
        {
            return this.DecryptWord(DecryptIt, "");
        }

        public string DecryptWord(string DecryptIt, string Weight)
        {
            string str3 = "";
            if (!this.IsValidUser)
            {
                throw new Exception("You can't use it without permission.");
            }
            if (Strings.Len(Strings.Trim(DecryptIt)) == 0)
            {
                return "";
            }
            long num2 = this.ReturnWeight(Weight);
            VBMath.Rnd(-1f);
            if (Strings.Len(Strings.Trim(Weight)) == 0)
            {
                VBMath.Randomize(5.0);
            }
            else
            {
                VBMath.Randomize((double)num2);
            }
            for (long i = 1L; i <= Strings.Len(DecryptIt); i += 1L)
            {
                long num = Strings.Asc(Strings.Mid(DecryptIt, (int)i, 1)) - ((int)Math.Round((double)((100f * VBMath.Rnd()) + 1f)));
                if (num > 0xffL)
                {
                    num -= 0xffL;
                }
                if (num < 1L)
                {
                    num += 0xffL;
                }
                str3 = str3 + StringType.FromChar(Strings.Chr((int)num));
            }
            return str3;
        }

        public string EncryptWord(string EncryptIt)
        {
            return this.EncryptWord(EncryptIt, "");
        }

        public string EncryptWord(string EncryptIt, string Weight)
        {
            string str3 = "";
            if (!this.IsValidUser)
            {
                throw new Exception("The encryption engine is only for KORMEE.");
            }
            if (Strings.Len(Strings.Trim(EncryptIt)) == 0)
            {
                return "";
            }
            long num2 = this.ReturnWeight(Weight);
            VBMath.Rnd(-1f);
            if (Strings.Len(Strings.Trim(Weight)) == 0)
            {
                VBMath.Randomize(5.0);
            }
            else
            {
                VBMath.Randomize((double)num2);
            }
            for (long i = 1L; i <= Strings.Len(EncryptIt); i += 1L)
            {
                long num = Strings.Asc(Strings.Mid(EncryptIt, (int)i, 1)) + ((int)Math.Round((double)((100f * VBMath.Rnd()) + 1f)));
                if (num > 0xffL)
                {
                    num -= 0xffL;
                }
                if (num < 1L)
                {
                    num += 0xffL;
                }
                str3 = str3 + StringType.FromChar(Strings.Chr((int)num));
            }
            return str3;
        }

        public void PopulateErrorString()
        {
            string[] strArray = new string[] { 
												 "Success.", "Invalid memory address.", "Unknown/invalid feature id option.", "Memory allocation failed.", "Too many open features.", "Feature access denied.", "Incompatible feature.", "HASP KEY not found.", "En-/decryption length too short.", "Invalid handle.", "Invalid file id / memory descriptor.", "Driver or support daemon version too old.", "Real time support not available.", "Generic error from host system call.", "Hardware key driver not found.", "Unrecognized info format.", 
												 "Request not supported.", "Invalid update object.", "Key with requested id was not found.", "Update data consistency check failed.", "Update not supported by this key.", "Update counter mismatch.", "Invalid vendor code.", "Requested encryption algorithm not supported.", "Invalid date / time.", "Clock out of power.", "Update requested ack., but no area to return it.", "Terminal services (remote terminal) detected.", "Feature type not implemented.", "Unknown algorithm.", "Signature check failed.", "Feature not found.", 
												 "Trace log not enabled."
											 };
            this.stringCollection = new StringCollection();
            this.stringCollection.AddRange(strArray);
            for (int i = this.stringCollection.Count; i <= 0x1f3; i++)
            {
                this.stringCollection.Insert(i, "");
            }
            strArray = new string[] { "Calling invalid object.", "A parameter is invalid.", "Already logged in.", "Already logged out." };
            this.stringCollection.AddRange(strArray);
            for (int j = this.stringCollection.Count; j <= 0x20c; j++)
            {
                this.stringCollection.Insert(j, "");
            }
            this.stringCollection.Insert(0x20d, "Unable to excecute/complete the operation.");
            for (int k = this.stringCollection.Count; k <= 0x257; k++)
            {
                this.stringCollection.Insert(k, "");
            }
            this.stringCollection.Insert(600, "No classic memory extension block available.");
            for (int m = this.stringCollection.Count; m <= 0x289; m++)
            {
                this.stringCollection.Insert(m, "");
            }
            this.stringCollection.Insert(650, "Invalid port type.");
            this.stringCollection.Insert(0x28b, "Invalid port.");
            for (int n = this.stringCollection.Count; n <= 0x2b9; n++)
            {
                this.stringCollection.Insert(n, "");
            }
            this.stringCollection.Insert(0x2ba, "Capability is not available.");
            this.stringCollection.Insert(0x2bb, "Internal API error.");
        }
        //Problem
        private long ReturnWeight(string Weight)
        {
            object obj2 = null;
            for (long i = 1L; i <= Strings.Len(Weight); i += 1L)
            {
                obj2 = ObjectType.AddObj(obj2, Strings.Asc(Strings.Mid(Weight, (int)i, 1)));
            }
            return LongType.FromObject(obj2);
        }


    }
}
