using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;
using MySql.Data.MySqlClient;
using System.Security.Cryptography;

namespace Transcripts_PDF_Generator
{
    class dbConnect
    {
      
        public MySqlConnection connection;
        public string server="localhost";
        private string database="student_portal";
        private string uid="root";
        private string passwword="";
        public string connectionString = "SERVER=localhost;" + "DATABASE=student_portal;" + "UID=root;" + "PASSWORD=;";
           
        public void dbconnection()
        {
            string connectionSTR = "SERVER=" + server + ";" + "DATABASE=" + database + ";" + "UID=" + uid + ";" + "PASSWORD=" + passwword+";";
            connection = new MySqlConnection(connectionSTR);

            connection.Open();
        }
        public string CalculateMD5Hash(string input)

{

    // step 1, calculate MD5 hash from input

    MD5 md5 = System.Security.Cryptography.MD5.Create();

    byte[] inputBytes = System.Text.Encoding.ASCII.GetBytes(input);

    byte[] hash = md5.ComputeHash(inputBytes);


    // step 2, convert byte array to hex string

    StringBuilder sb = new StringBuilder();

    for (int i = 0; i < hash.Length; i++)

    {

        sb.Append(hash[i].ToString("X2"));

    }

    return sb.ToString();

}
        public string GetFullEncryption(string input)
        {

            char[] array = input.ToCharArray();
            for (int i = 0; i < array.Length; i++)
            {
                char let = array[i];
                if (let == 'a')
                    array[i] = '!';
                else if (let == 'b')
                    array[i] = '@';
                else if (let == 'c')
                    array[i] = '#';
                else if (let == 'd')
                    array[i] = '$';
                else if (let == 'e')
                    array[i] = '%';
                else if (let == 'f')
                    array[i] = '^';
                else if (let == 'g')
                    array[i] = '&';
                else if (let == 'h')
                    array[i] = '*';
                else if (let == 'i')
                    array[i] = '(';
                else if (let == 'j')
                    array[i] = ')';
                else if (let == 'k')
                    array[i] = '-';
                else if (let == 'l')
                    array[i] = '_';
                else if (let == 'm')
                    array[i] = '=';
                else if (let == 'n')
                    array[i] = '+';
                else if (let == 'o')
                    array[i] = '\\';
                else if (let == 'p')
                    array[i] = '|';
                else if (let == 'q')
                    array[i] = '/';
                else if (let == 'r')
                    array[i] = '>';
                else if (let == 's')
                    array[i] = '<';
                else if (let == 't')
                    array[i] = '?';
                else if (let == 'u')
                    array[i] = '[';
                else if (let == 'v')
                    array[i] = ']';
                else if (let == 'w')
                    array[i] = '~';
                else if (let == 'x')
                    array[i] = '{';
                else if (let == 'y')
                    array[i] = '}';
                else if (let == '0')
                    array[i] = 'Z';
                else if (let == '1')
                    array[i] = 'Y';
                else if (let == '2')
                    array[i] = 'X';
                else if (let == '3')
                    array[i] = 'W';
                else if (let == '4')
                    array[i] = 'V';
                else if (let == '5')
                    array[i] = 'U';
                else if (let == '6')
                    array[i] = 'T';
                else if (let == '7')
                    array[i] = 'S';
                else if (let == '8')
                    array[i] = 'R';
                else if (let == '9')
                    array[i] = 'Q';
                else if (let == 'A')
                    array[i] = 'q';
                else if (let == 'B')
                    array[i] = 'w';
                else if (let == 'C')
                    array[i] = 'e';
                else if (let == 'D')
                    array[i] = 'r';
                else if (let == 'E')
                    array[i] = 't';
                else if (let == 'F')
                    array[i] = 'y';
                else if (let == 'G')
                    array[i] = 'u';
                else if (let == 'H')
                    array[i] = 'i';
                else if (let == 'I')
                    array[i] = 'o';
                else if (let == 'J')
                    array[i] = 'p';
                else if (let == 'K')
                    array[i] = 'a';
                else if (let == 'L')
                    array[i] = 's';
                else if (let == 'M')
                    array[i] = 'd';
                else if (let == 'N')
                    array[i] = 'f';
                else if (let == 'O')
                    array[i] = 'g';
                else if (let == 'P')
                    array[i] = 'h';
                else if (let == 'Q')
                    array[i] = 'j';
                else if (let == 'R')
                    array[i] = 'k';
                else if (let == 'S')
                    array[i] = 'l';
                else if (let == 'T')
                    array[i] = 'z';
                else if (let == 'U')
                    array[i] = 'x';
                else if (let == 'V')
                    array[i] = 'c';
                else if (let == 'W')
                    array[i] = 'v';
                else if (let == 'X')
                    array[i] = 'b';
                else if (let == 'Y')
                    array[i] = 'n';
                else if (let == 'Z')
                    array[i] = 'm';
                else if (let == '~')
                    array[i] = 'P';
                else if (let == '`')
                    array[i] = 'O';
                else if (let == '!')
                    array[i] = 'N';
                else if (let == '@')
                    array[i] = 'M';
                else if (let == '#')
                    array[i] = 'L';
                else if (let == '$')
                    array[i] = 'K';
                else if (let == '%')
                    array[i] = 'J';
                else if (let == '^')
                    array[i] = 'I';
                else if (let == '&')
                    array[i] = 'H';
                else if (let == '*')
                    array[i] = 'G';
                else if (let == '(')
                    array[i] = 'F';
                else if (let == ')')
                    array[i] = 'E';
                else if (let == '-')
                    array[i] = 'D';
                else if (let == '_')
                    array[i] = 'C';
                else if (let == '+')
                    array[i] = 'B';
                else if (let == '=')
                    array[i] = 'A';
                else if (let == '{')
                    array[i] = '0';
                else if (let == '[')
                    array[i] = '1';
                else if (let == '}')
                    array[i] = '2';
                else if (let == ']')
                    array[i] = '3';
                else if (let == '|')
                    array[i] = '4';
                else if (let == '\\')
                    array[i] = '5';
                else if (let == ':')
                    array[i] = '6';
                else if (let == ';')
                    array[i] = '7';
                else if (let == '"')
                    array[i] = '8';
                else if (let == '\'')
                    array[i] = '9';
                else if (let == '<')
                    array[i] = ':';
                else if (let == ',')
                    array[i] = ';';
                else if (let == '>')
                    array[i] = '"';
                else if (let == '.')
                    array[i] = '`';
                else if (let == '?')
                    array[i] = '.';
                else if (let == '/')
                    array[i] = '\'';
            }
            return new string(array);

        }

    }
}
