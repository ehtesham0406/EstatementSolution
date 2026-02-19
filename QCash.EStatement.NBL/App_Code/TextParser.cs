using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace QCash.EStatement.NBL.App_Code
{
   public class TextParser
    {
       public TextObj parser(string[] txt)
       {
           TextObj _TextObj =new TextObj();
           try
           {
               for (int i=0; i <= 11;i++ )
               {
                   if (txt[i].Trim() != "()")
                   {
                       #region
                       switch (i)
                       {
                           case 0:
                               {
                                   _TextObj.IDClient = txt[i].Trim();
                                   break;
                               }
                           case 1:
                               {
                                   _TextObj.PAN = txt[i].Trim();
                                   break;
                               }
                           case 2:
                               {
                                   _TextObj.SDate = txt[i].Trim();
                                   break;
                               }
                           case 3:
                               {
                                   _TextObj.Branch = txt[i].Trim();
                                   break;
                               }
                           case 4:
                               {
                                   _TextObj.AmountLimit = txt[i].Trim();
                                   break;
                               }
                           case 5:
                               {
                                   _TextObj.Client = txt[i].Trim();
                                   break;
                               }
                           case 6:
                               {
                                   _TextObj.CardType = txt[i].Trim();
                                   break;
                               }
                           case 7:
                               {
                                   _TextObj.Code = txt[i].Trim();
                                   break;
                               }
                           case 8:
                               {
                                   _TextObj.Address1 = txt[i].Trim();
                                   break;
                               }
                           case 9:
                               {
                                   _TextObj.Address2 = txt[i].Trim();
                                   break;
                               }
                           case 10:
                               {
                                   _TextObj.Country = txt[i].Trim();
                                   break;
                               }
                           case 11:
                               {
                                   _TextObj.Mobile = txt[i].Trim();
                                   break;
                               }
                       }

                       #endregion

                   }
                   else
                       break;

               }
           
           }
           catch(Exception ex) 
           { 
               
           }

           return _TextObj;
       }

       public string[] arryParse(string[] parsearry)
       {
           string[] _arryParse;
           if (parsearry.Length > 13)
           {
               _arryParse = new string[parsearry.Length - 14];
               for (int i = 0; i <= parsearry.Length - 16; i++)
               {
                   _arryParse[i] = parsearry[14 + i];

               }
           }
           else
           {
               _arryParse = null;
           }
          
           return _arryParse;
       
       }

     

    }
}
