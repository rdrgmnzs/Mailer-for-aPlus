using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Settings
{
    class PasswordEncoder
    {

        // Change the two values below to be something other than the example.
        // Once changed and in use, do not change the value below again or you
        // won't be able to decrypt previously stored passwords.
        private const string mByteArray = "%$5>#%2(2sUat#l)URa0$!@";

        static byte[] mInitializationVector = { 0x08, 0x54, 0xA6, 0x78, 0x98, 0xAF, 0xF3, 0xEB };

        public static string EncryptWithByteArray(string inPassword)
        {
            string mEncryptedPassword = EncryptWithByteArray(inPassword, mByteArray);
            return mEncryptedPassword;
        }

        private static string EncryptWithByteArray(string inPassword, string inByteArray)
        {
            try
            {
                byte[] tmpKey = new byte[20];
                tmpKey = System.Text.Encoding.UTF8.GetBytes(inByteArray.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                byte[] inputArray = System.Text.Encoding.UTF8.GetBytes(inPassword);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(tmpKey, mInitializationVector), CryptoStreamMode.Write);
                cs.Write(inputArray, 0, inputArray.Length);
                cs.FlushFinalBlock();
                return Convert.ToBase64String(ms.ToArray());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static string DecryptWithByteArray(string mEncryptedPassword)
        {
            return DecryptWithByteArray(mEncryptedPassword, mByteArray);
        }

        private static string DecryptWithByteArray(string strText, string strEncrypt)
        {
           try
           {
                byte[] tmpKey = new byte[20];
                tmpKey = System.Text.Encoding.UTF8.GetBytes(strEncrypt.Substring(0, 8));
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                Byte[] inputByteArray = inputByteArray = Convert.FromBase64String(strText);
                MemoryStream ms = new MemoryStream();
                CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(tmpKey, mInitializationVector), CryptoStreamMode.Write);
                cs.Write(inputByteArray, 0, inputByteArray.Length);
                cs.FlushFinalBlock();
                System.Text.Encoding encoding = System.Text.Encoding.UTF8;
                return encoding.GetString(ms.ToArray());
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string ByteArray
        {
            get { return mByteArray; }
        }
    }
}
