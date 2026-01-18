using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.Remoting.Lifetime;
using System.Text;
using System.Threading.Tasks;

namespace AlgoÖdev2
{
    internal class Program

    {




        // defining global variables 

        static string[,] excelTable = new string[15, 10];// Row,Column
        static string[,] dataTypes = new string[15, 10];
        static int currentColSize = 5;
        static int currentRowSize = 8;

        // defining helper funstions
        static void InitArrays()
        {

            for (int i = 0; i < 15; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    dataTypes[i, j] = "unassigned";

                    if (excelTable[i, j] == null)
                    {
                        excelTable[i, j] = "";
                    }
                }
            }
        }
        static bool IntegerCheck(string val)
        {
            if (val == "" || val == null) return false;

            int startIndex = 0;

            // control for negative numbers
            if (val[0] == '-')
            {
                if (val.Length == 1) return false;
                startIndex = 1;
            }

            for (int i = startIndex; i < val.Length; i++)
            {
                int asciiVal = (int)val[i];


                if (asciiVal < 48 || asciiVal > 57)
                {
                    return false;
                }
            }

            return true;
        }
        static int[] decodeCell(string cellName)
        {
            string charvalue = cellName.Substring(0, 1);
            string number = cellName.Substring(1);
            int indChar = (int)charvalue[0];
            indChar = indChar - 65; // Column Index 
            int num = Convert.ToInt32(number);
            num = num - 1;  // Row Index 

            int[] ans = { num, indChar };

            return ans;
        }
        static string removinglettersint(string s1, char s2)
        {
            string ans = "";
            for (int i = 0; i < s1.Length; i++)
            {
                if (s1[i] != s2)
                {
                    ans += s1[i];
                }
            }
            return ans;

        }
        static string removinglettersstring(string s1, string s2)
        {
            string ans = "";
            string longstr = "";
            string shortst = "";
            int i = 0;
            if (s1.Length > s2.Length)
            {
                longstr = s1;
                shortst = s2;

            }
            else if (s2.Length > s1.Length)
            {
                longstr = s2;
                shortst = s1;
            }

            while (i < longstr.Length)
            {
                // control check for substring
                if (i <= (longstr.Length - shortst.Length))
                {
                    string piece = longstr.Substring(i, shortst.Length);
                    if (piece == shortst)
                    {
                        i += shortst.Length;
                        continue;
                    }
                }

                ans += longstr[i];
                i++;
            }
            return ans;
        }

        static void readdata()
        {
            string path = "spreadsheet.txt";

            if (File.Exists(path))
            {
                StreamReader f = File.OpenText(path);
                // read first line
                if (!f.EndOfStream)
                {
                    string firstLine = f.ReadLine();
                    string[] size = firstLine.Split(';');

                    if (size.Length >= 2)
                    {
                        currentRowSize = Convert.ToInt32(size[0]);
                        currentColSize = Convert.ToInt32(size[1]);
                    }
                }

                // read data
                string line;
                while (!f.EndOfStream)
                {
                    line = f.ReadLine();
                    if (line == "") continue;

                    string[] parcalar = line.Split(';');

                    if (parcalar.Length >= 4)
                    {
                        int r = Convert.ToInt32(parcalar[0]);
                        int c = Convert.ToInt32(parcalar[1]);
                        string veri = parcalar[2];
                        string dataType = parcalar[3];


                        excelTable[r, c] = veri;
                        dataTypes[r, c] = dataType;
                    }
                }
                f.Close();
            }
        }

        static void writedata()
        {
            string path = "spreadsheet.txt";
            StreamWriter f = File.CreateText(path);


            // first line include currentRowSize and CurrentCol size 
            f.WriteLine(currentRowSize + ";" + currentColSize);


            for (int r = 0; r < currentRowSize; r++)
            {
                for (int c = 0; c < currentColSize; c++)
                {
                    if (excelTable[r, c] != null && excelTable[r, c] != "")
                    {
                        string line = r + ";" + c + ";" + excelTable[r, c] + ";" + dataTypes[r, c];
                        f.WriteLine(line);
                    }
                }
            }
            f.Close();
        }

        // defining all main and helper functions
        static void ClearAll()
        {
            for (int i = 0; i < excelTable.GetLength(0); i++)
            {
                for (int j = 0; j < excelTable.GetLength(1); j++)
                {
                    excelTable[i, j] = "";
                    dataTypes[i, j] = "unassigned";
                }
            }
        }

        static void ClearCell(string cellPosition)
        {
            int[] places = decodeCell(cellPosition);
            excelTable[places[0], places[1]] = "";
            dataTypes[places[0], places[1]] = "unassigned";
        }

        static void AssignValue(string place, string dataType, string value)
        {
            int[] places = decodeCell(place);
            // places[0] -> Row (Satır)
            // places[1] -> Col (Sütun)

            // error check for bounds
            if (places[0] >= 15 || places[1] >= 10)
            {
                Console.WriteLine("The specified row or column exceeds the maximum grid size!");
                Console.WriteLine("");
                return;
            }
            if (places[0] > currentRowSize - 1 || places[1] > currentColSize - 1)
            {
                Console.WriteLine("The specified row or column exceeds then current grid size!");
                Console.WriteLine("");
                return;

            }

            if (dataType == "integer")
            {
                bool isNumber = IntegerCheck(value);
                if (isNumber == false)
                {
                    Console.WriteLine("Data Type Mismatch! You cannot assign a non-integer value to an integer cell.");
                    Console.WriteLine("");
                    return;
                }
            }
            if (dataType == "string")
            {
                bool isNumber = IntegerCheck(value);
                if (isNumber == true)
                {
                    Console.WriteLine("Data Type Mismatch! You cannot assign a non-string value to an string cell.");
                    Console.WriteLine("");
                    return;
                }
            }

            // asssigment operation
            if (places[0] < 15 && places[1] < 10)
            {

                excelTable[places[0], places[1]] = value;

                dataTypes[places[0], places[1]] = dataType;


            }
            Console.WriteLine("Operation is done!");
            Console.WriteLine();
            // buraya hata mesajı eklenecek unutma 

        }
        static void PrintTable()
        {
            // print column names
            Console.Write("    ");
            for (int c = 0; c < currentColSize; c++)
            {
                char colName = (char)('A' + c);
                Console.Write($"   {colName}    ");
            }
            Console.WriteLine();
            Console.WriteLine("    " + new string('-', currentColSize * 8));

            // print rows and datas
            for (int r = 0; r < currentRowSize; r++)
            {
                // print | and space 3
                Console.Write($"{r + 1,3} |");

                for (int c = 0; c < currentColSize; c++)
                {
                    string rawValue = excelTable[r, c];

                    // error check for null strings on the table
                    if (rawValue == null) rawValue = "";

                    string displayValue = "";

                    // length check for strings
                    if (rawValue.Length > 5)
                    {
                        displayValue = rawValue.Substring(0, 5) + "_";
                    }
                    else
                    {
                        displayValue = rawValue;
                    }

                    Console.Write($" {displayValue,-6}|");
                }
                Console.WriteLine();
            }
        }

        static void Copy(string c1, string c2)
        {
            int[] c1places = decodeCell(c1);
            int[] c2places = decodeCell(c2);
            if (c1places[0] > currentRowSize - 1 || c1places[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (c2places[0] > currentRowSize - 1 || c2places[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }

            excelTable[c2places[0], c2places[1]] = excelTable[c1places[0], c1places[1]];
            dataTypes[c2places[0], c2places[1]] = dataTypes[c1places[0], c1places[1]];

            Console.WriteLine("Operation is done");
        }
        static void CopyColumn(string c1, string c2)
        {
            int indc1 = c1[0] - 65;
            int indc2 = c2[0] - 65;
            if ((indc1 > (currentColSize - 1) || indc1 < 0) || ((indc2 > currentColSize - 1) || indc2 < 0))
            {
                Console.WriteLine("Illegal position assignment!");
                return;
            }

            for (int i = 0; i < currentRowSize; i++)
            {
                excelTable[i, indc2] = excelTable[i, indc1];
                dataTypes[i, indc2] = dataTypes[i, indc1];
            }
            Console.WriteLine("Operation is done");

        }
        static void CopyRow(string r1, string r2)
        {
            if (!IntegerCheck(r1) || !IntegerCheck(r2))
            {
                Console.WriteLine("invalid input for this operation!");
                return;
            }
            int indr1 = Convert.ToInt32(r1) - 1;
            int indr2 = Convert.ToInt32(r2) - 1;

            if ((indr1 > (currentRowSize - 1) || indr1 < 0) || ((indr2 > currentRowSize - 1) || indr2 < 0))
            {
                Console.WriteLine("Illegal position assignment!");
                return;
            }
            for (int i = 0; i < currentColSize; i++)
            {
                excelTable[indr2, i] = excelTable[indr1, i];
                dataTypes[indr2, i] = dataTypes[indr1, i];
            }
            Console.WriteLine("Operation is done");


        }


        static void AddRow(string rowInd, string direction)
        {

            rowInd = rowInd.Trim();
            direction = direction.Trim();
            int inputRow = Convert.ToInt32(rowInd);
            int rowIndex = inputRow - 1;

            // check for out of bounds 
            if (inputRow >= 15 || inputRow < 1)
            {
                Console.WriteLine("Out of bounds exception!");
                return;
            }

            if (currentRowSize >= 15)
            {
                Console.WriteLine("Out of bounds exception!");
                return;
            }

            //check for availability 
            if (inputRow > currentRowSize)
            {
                Console.WriteLine("Illegal position assignment!");
                return;
            }



            int targetIndex = -1;

            // remove spaces from direction
            direction = direction.Trim();

            // select the target according to the direction
            if (direction == "up")
            {
                targetIndex = rowIndex;
            }
            else if (direction == "down")
            {
                targetIndex = rowIndex + 1;
            }
            else
            {

                Console.WriteLine("Syntax error!");
                return;
            }

            // shifting
            for (int r = currentRowSize - 1; r >= targetIndex; r--)
            {
                for (int c = 0; c < currentColSize; c++)
                {
                    excelTable[r + 1, c] = excelTable[r, c];
                    dataTypes[r + 1, c] = dataTypes[r, c];
                }
            }

            // clear new line
            for (int c = 0; c < currentColSize; c++)
            {
                excelTable[targetIndex, c] = "";
                dataTypes[targetIndex, c] = "unassigned";
            }

            currentRowSize++;
            Console.WriteLine("Operation is done!");

        }

        static void AddColumn(string colId, string direction)
        {
            colId = colId.Trim();
            direction = direction.Trim();
            string charvalue = colId.Substring(0, 1);
            int indChar = (int)charvalue[0];
            indChar = indChar - 65;
            // check for bounds
            if (indChar >= 10 || indChar < 0)
            {
                Console.WriteLine("Out of bounds exception!");
                return;

            }
            if (currentColSize >= 10)
            {
                Console.WriteLine("Out of bounds exception!");
                return;
            }
            if (indChar > currentColSize - 1)
            {
                Console.WriteLine("Illegal position assignment!");
                Console.WriteLine("");
                return;
            }
            // define target index according to direction
            int targetIndex = -1;
            if (direction == "left")
            {
                targetIndex = indChar;
            }
            else if (direction == "right")
            {
                targetIndex = indChar + 1;
            }

            for (int c = currentColSize - 1; c >= targetIndex; c--)
            {
                for (int r = 0; r < currentRowSize; r++)
                {
                    excelTable[r, c + 1] = excelTable[r, c];
                    dataTypes[r, c + 1] = dataTypes[r, c];

                }




            }
            // assignment operations 
            for (int t = 0; t < currentRowSize; t++)
            {
                excelTable[t, targetIndex] = "";
                dataTypes[t, targetIndex] = "unassigned";
            }

            currentColSize += 1;
            Console.WriteLine("Operation is done!");
            Console.WriteLine("");


        }

        static void X(string c1, string c2)
        {
            c1 = c1.Trim();
            c2 = c2.Trim();
            int c1col_ind = c1[0] - 'A';
            int c2col_ind = c2[0] - 'A';

            int c1row_ind = Convert.ToInt32(c1.Substring(1)) - 1;
            int c2row_ind = Convert.ToInt32(c2.Substring(1)) - 1;

            // check for bounds
            if (c1row_ind > currentRowSize - 1 || c1col_ind > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (c2row_ind > currentRowSize - 1 || c2col_ind > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }



            excelTable[c2row_ind, c2col_ind] = excelTable[c1row_ind, c1col_ind];
            excelTable[c1row_ind, c1col_ind] = "";

            dataTypes[c2row_ind, c2col_ind] = dataTypes[c1row_ind, c1col_ind];
            dataTypes[c1row_ind, c1col_ind] = "unassigned";

            Console.WriteLine("Operation is done!");

        }

        static void XColumn(string c1, string c2)
        {
            c1 = c1.Trim();
            c2 = c2.Trim();
            string c1v = c1.Substring(0, 1);
            string c2v = c2.Substring(0, 1);
            int c1ind = c1[0] - 65;
            int c2ind = c2[0] - 65;

            if ((c1ind > 9 || c1ind < 0) || (c2ind > 9 || c2ind < 0))
            {
                Console.WriteLine("Out of bounds exception!");
                return;

            }
            if ((c1ind > (currentColSize - 1)) || (c2ind > (currentColSize - 1)))
            {
                Console.WriteLine("Illegal position assignment!");
                return;

            }

            for (int r = 0; r < currentRowSize; r++)
            {
                excelTable[r, c2ind] = excelTable[r, c1ind];
                dataTypes[r, c2ind] = dataTypes[r, c1ind];
            }

            //clear operations

            for (int cl = 0; cl < currentRowSize; cl++)
            {
                excelTable[cl, c1ind] = "";
                dataTypes[cl, c1ind] = "unassigned";
            }
            Console.WriteLine("Operation is done!");

        }

        static void XRow(string r1, string r2)
        {
            r1 = r1.Trim();
            r2 = r2.Trim();

            int r1ind = Convert.ToInt32(r1) - 1;
            int r2ind = Convert.ToInt32(r2) - 1;

            if ((r1ind > 14 || r1ind < 0) || (r2ind > 14) || (r2ind < 0))
            {
                Console.WriteLine("Out of bounds exception!");
                return;

            }
            if ((r1ind > (currentRowSize - 1)) || (r2ind > (currentRowSize - 1)))
            {
                Console.WriteLine("Illegal position assignment!");
                return;

            }

            // copy operation

            for (int c = 0; c < currentColSize; c++)
            {
                excelTable[r2ind, c] = excelTable[r1ind, c];
                dataTypes[r2ind, c] = dataTypes[r1ind, c];
            }

            // cut operation

            for (int c = 0; c < currentColSize; c++)
            {
                excelTable[r1ind, c] = "";
                dataTypes[r1ind, c] = "unassigned";
            }

            Console.WriteLine("Operation is done!");

        }

        static string ReverseString(string str)
        {
            string ans = "";
            for (int i = str.Length - 1; i >= 0; i--)
            {
                ans += str[i];
            }
            return ans;
        }
        static void mult(string c1, string c2, string outputcell)
        {


            int[] cell1 = decodeCell(c1);
            int[] cell2 = decodeCell(c2);
            int[] cellout = decodeCell(outputcell);
            int ans = 0;
            string ansstr = "";

            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }



            string typec1 = dataTypes[cell1[0], cell1[1]];
            string typec2 = dataTypes[cell2[0], cell2[1]];

            if (dataTypes[cell1[0], cell1[1]] == "unassigned" || dataTypes[cell2[0], cell2[1]] == "unassigned")
            {
                Console.WriteLine("Illegal operation! One of the input cells is unassigned.");
                return;
            }


            if (typec1 == typec2)
            {
                if (typec1 == "string")
                {
                    Console.WriteLine("Illegal operation!");
                    Console.WriteLine("String – String operation is not allowed!");
                    return;
                }
                if (typec1 == "integer")
                {
                    int value1 = Convert.ToInt32(excelTable[cell1[0], cell1[1]]);
                    int value2 = Convert.ToInt32(excelTable[cell2[0], cell2[1]]);
                    ans = (value1 * value2);

                    excelTable[cellout[0], cellout[1]] = Convert.ToString(ans);
                    dataTypes[cellout[0], cellout[1]] = "integer";
                    Console.WriteLine("Operation is done");
                }

            }
            if (typec1 != typec2)
            {
                if (typec1 == "string")
                {
                    string value1 = excelTable[cell1[0], cell1[1]];
                    int value2 = Convert.ToInt32(excelTable[cell2[0], cell2[1]]);
                    if (value2 > 0)
                    {
                        for (int i = 0; i < value2; i++)
                        {
                            ansstr += value1;
                        }
                    }
                    if (value2 < 0)
                    {
                        for (int i = 0; i < -1 * value2; i++)
                        {
                            ansstr += ReverseString(value1);
                        }
                    }

                }
                if (typec2 == "string")
                {
                    int value1 = Convert.ToInt32(excelTable[cell1[0], cell1[1]]);
                    string value2 = excelTable[cell2[0], cell2[1]];
                    if (value1 > 0)
                    {
                        for (int i = 0; i < value1; i++)
                        {
                            ansstr += value2;
                        }
                    }
                    if (value1 < 0)
                    {
                        for (int i = 0; i < -1 * value1; i++)
                        {
                            ansstr += ReverseString(value2);


                        }

                    }
                }

                excelTable[cellout[0], cellout[1]] = ansstr;
                dataTypes[cellout[0], cellout[1]] = "string";
                Console.WriteLine("Operation is done!");
            }


        }

        static void sum3param(string c1, string c2, string outputcell)
        {
            int[] cell1 = decodeCell(c1);
            int[] cell2 = decodeCell(c2);
            int[] cellout = decodeCell(outputcell);





            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c1 + " is out of the bounds!");
                return;

            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c2 + " is out of the bounds!");
                return;

            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(outputcell + " is out of the bounds!");
                return;

            }
            string val1 = excelTable[cell1[0], cell1[1]];
            string val2 = excelTable[cell2[0], cell2[1]];

            //datatype check
            if (dataTypes[cell1[0], cell1[1]] == "unassigned" || dataTypes[cell2[0], cell2[1]] == "unassigned")
            {
                Console.WriteLine("Illegal operation! One of the input cells is unassigned.");
                return;
            }

            if ((dataTypes[cell1[0], cell1[1]] == "string") || (dataTypes[cell2[0], cell2[1]] == "string"))
            {
                Console.WriteLine("Concatenation operation will be applied!");
                Console.Write("Please select letter case :");
                string option = Console.ReadLine();
                string ansconcatenation = "";
                Console.WriteLine("");
                if (option == "up")
                {
                    val1 = val1.ToUpper();
                    val2 = val2.ToUpper();

                }
                else if (option == "down")
                {
                    val1 = val1.ToLower();
                    val2 = val2.ToLower();
                }
                ansconcatenation = val1 + val2;

                excelTable[cellout[0], cellout[1]] = ansconcatenation;
                dataTypes[cellout[0], cellout[1]] = "string";
                Console.WriteLine("Operation is done!");
                Console.Write(c1 + ": " + val1);
                Console.WriteLine("");
                Console.Write(c2 + ": " + val2);
                Console.WriteLine("");
                Console.Write("Output: " + ansconcatenation);
                Console.WriteLine("");
            }
            else if ((dataTypes[cell1[0], cell1[1]] == "integer") && (dataTypes[cell2[0], cell2[1]] == "integer"))
            {
                int ansaddition;
                int var1 = Convert.ToInt32(excelTable[cell1[0], cell1[1]]);
                int var2 = Convert.ToInt32(excelTable[cell2[0], cell2[1]]);

                ansaddition = var1 + var2;

                Console.WriteLine("Operation is done!");
                Console.Write(c1 + ": " + var1);
                Console.WriteLine("");
                Console.Write(c2 + ": " + var2);
                Console.WriteLine("");
                Console.Write("Output: " + ansaddition);
                Console.WriteLine("");
                excelTable[cellout[0], cellout[1]] = Convert.ToString(ansaddition);
                dataTypes[cellout[0], cellout[1]] = "integer";
            }

        }
        static void sum4param(string c1, string c2, string c3, string outputcell)
        {
            int[] cell1 = decodeCell(c1);
            int[] cell2 = decodeCell(c2);
            int[] cell3 = decodeCell(c3);
            int[] cellout = decodeCell(outputcell);


            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c1 + " is out of the bounds!");
                return;

            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c2 + " is out of the bounds!");
                return;

            }
            if (cell3[0] > currentRowSize - 1 || cell3[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c3 + " is out of the bounds!");
                return;

            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(outputcell + " is out of the bounds!");
                return;

            }
            if (dataTypes[cell1[0], cell1[1]] == "unassigned" || dataTypes[cell2[0], cell2[1]] == "unassigned" || dataTypes[cell3[0], cell3[1]] == "unassigned")
            {
                Console.WriteLine("Illegal operation! One of the input cells is unassigned.");
                return;
            }
            string val1 = excelTable[cell1[0], cell1[1]];
            string val2 = excelTable[cell2[0], cell2[1]];
            string val3 = excelTable[cell3[0], cell3[1]];

            if ((dataTypes[cell1[0], cell1[1]] == "string") || (dataTypes[cell2[0], cell2[1]] == "string") || (dataTypes[cell3[0], cell3[1]] == "string"))
            {
                Console.WriteLine("Concatenation operation will be applied!");
                Console.Write("Please select letter case :");
                string option = Console.ReadLine();
                string ansconcatenation = "";
                Console.WriteLine("");

                if (option == "up")
                {
                    val1 = val1.ToUpper();
                    val2 = val2.ToUpper();
                    val3 = val3.ToUpper();
                }
                else if (option == "down")
                {
                    val1 = val1.ToLower();
                    val2 = val2.ToLower();
                    val3 = val3.ToLower();
                }
                ansconcatenation = val1 + val2 + val3;

                excelTable[cellout[0], cellout[1]] = ansconcatenation;
                dataTypes[cellout[0], cellout[1]] = "string";
                Console.WriteLine("Operation is done!");
                Console.Write(c1 + ": " + val1);
                Console.WriteLine("");
                Console.Write(c2 + ": " + val2);
                Console.WriteLine("");
                Console.Write(c3 + ": " + val3);
                Console.WriteLine("");
                Console.Write("Output: " + ansconcatenation);
                Console.WriteLine("");


            }
            else if ((dataTypes[cell1[0], cell1[1]] == "integer") && (dataTypes[cell2[0], cell2[1]] == "integer") && (dataTypes[cell3[0], cell3[1]] == "integer"))
            {
                int ansaddition;
                int var1 = Convert.ToInt32(excelTable[cell1[0], cell1[1]]);
                int var2 = Convert.ToInt32(excelTable[cell2[0], cell2[1]]);
                int var3 = Convert.ToInt32(excelTable[cell3[0], cell3[1]]);
                ansaddition = var1 + var2 + var3;

                Console.WriteLine("Operation is done!");
                Console.Write(c1 + ": " + var1);
                Console.WriteLine("");
                Console.Write(c2 + ": " + var2);
                Console.WriteLine("");
                Console.Write(c3 + ": " + var3);
                Console.WriteLine("");
                Console.Write("Output: " + ansaddition);
                Console.WriteLine("");
                excelTable[cellout[0], cellout[1]] = Convert.ToString(ansaddition);
                dataTypes[cellout[0], cellout[1]] = "integer";
            }


        }

        static void divide(string c1, string c2, string outputcell)
        {
            int[] cell1 = decodeCell(c1);
            int[] cell2 = decodeCell(c2);
            int[] cellout = decodeCell(outputcell);
            int ans = 0;
            string ansstr = "";
            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            string val1 = excelTable[cell1[0], cell1[1]];
            string val2 = excelTable[cell2[0], cell2[1]];
            string type1 = dataTypes[cell1[0], cell1[1]];
            string type2 = dataTypes[cell2[0], cell2[1]];
            string text = "";
            int num = 0;
            if (dataTypes[cell1[0], cell1[1]] == "unassigned" || dataTypes[cell2[0], cell2[1]] == "unassigned")
            {
                Console.WriteLine("Illegal operation! One of the input cells is unassigned.");
                return;
            }

            // bound check
            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c1 + " is out of the bounds!");
                return;

            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(c2 + " is out of the bounds!");
                return;

            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("Illegal operation!");
                Console.WriteLine(outputcell + " is out of the bounds!");
                return;

            }



            if (type1 == type2)
            {
                if (type1 == "integer")
                {
                    if (Convert.ToInt32(val2) == 0)
                    {
                        Console.WriteLine("you cannot divide by zero");
                        return;
                    }
                    ans = Convert.ToInt32(val1) / Convert.ToInt32(val2);
                    excelTable[cellout[0], cellout[1]] = Convert.ToString(ans);
                    dataTypes[cellout[0], cellout[1]] = "integer";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + ans);
                }
                else
                {
                    Console.WriteLine("Illegal operation!");
                    Console.WriteLine("String – String operation is not allowed!");
                    return;
                }
            }
            if (type1 != type2)
            {
                int c = 0;
                if (type1 == "integer" && type2 == "string")
                {
                    text = val2;
                    num = Convert.ToInt32(val1);
                }
                else if (type1 == "string" && type2 == "integer")
                {
                    text = val1;
                    num = Convert.ToInt32(val2);
                }
                if (num == 0)
                {
                    Console.WriteLine("you cannot divide by zero");
                    return;
                }

                if (num > 0)
                {
                    c = text.Length / num;
                    ansstr = text.Substring(0, c);
                    excelTable[cellout[0], cellout[1]] = ansstr;
                    dataTypes[cellout[0], cellout[1]] = "string";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + ansstr);

                }
                if (num < 0)
                {
                    c = text.Length / Math.Abs(num);
                    string temp = ReverseString(text);
                    ansstr = temp.Substring(0, c);
                    excelTable[cellout[0], cellout[1]] = ansstr;
                    dataTypes[cellout[0], cellout[1]] = "string";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + ansstr);
                }



            }
        }

        static void hash(string c1, string c2, string outputcell)
        {
            int[] cell1 = decodeCell(c1);
            int[] cell2 = decodeCell(c2);
            int[] cellout = decodeCell(outputcell);
            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            string ansstr = "";
            string val1 = excelTable[cell1[0], cell1[1]];
            string val2 = excelTable[cell2[0], cell2[1]];
            string type1 = dataTypes[cell1[0], cell1[1]];
            string type2 = dataTypes[cell2[0], cell2[1]];
            int num = 0;
            string text = "";

            if (dataTypes[cell1[0], cell1[1]] == "unassigned" || dataTypes[cell2[0], cell2[1]] == "unassigned")
            {
                Console.WriteLine("Illegal operation! One of the input cells is unassigned.");
                return;
            }

            if (type1 == type2)
            {
                if (type1 == "integer")
                {
                    Console.WriteLine("Illegal operation!");
                    Console.WriteLine("Integer – Integer operation is not allowed! ");
                    return;
                }
                if (type2 == "string")
                {
                    Console.WriteLine("Illegal operation!");
                    Console.WriteLine("String – String operation is not allowed!");
                    return;
                }

            }
            if (type1 != type2)
            {
                if (type1 == "string" && type2 == "integer")
                {
                    text = val1;

                    num = Convert.ToInt32(val2);

                    if (num < -20 || num > 30)
                    {
                        Console.WriteLine("Illegal operation!");
                        Console.WriteLine("The integer parameter should be in [-20, 30]");
                        return;
                    }
                }
                if (type1 == "integer" && type2 == "string")
                {
                    text = val2;
                    num = Convert.ToInt32(val1);
                    if (num < -20 || num > 30)
                    {
                        Console.WriteLine("Illegal operation!");
                        Console.WriteLine("The integer parameter should be in [-20, 30]");
                        return;
                    }
                }
                for (int i = 0; i < text.Length; i++)
                {
                    int normalascii = (int)text[i];
                    int hashedascii = normalascii + num;
                    ansstr += (char)hashedascii;
                }
                excelTable[cellout[0], cellout[1]] = ansstr;
                dataTypes[cellout[0], cellout[1]] = "string";
                Console.WriteLine("Operation is done!");
                Console.WriteLine(c1 + ": " + val1);
                Console.WriteLine(c2 + ": " + val2);
                Console.WriteLine("Encrypted value: " + ansstr);

            }

        }
        static void substract(string c1, string c2, string outputcell)
        {
            int[] cell1 = decodeCell(c1);
            int[] cell2 = decodeCell(c2);
            int[] cellout = decodeCell(outputcell);
            if (cell1[0] > currentRowSize - 1 || cell1[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cell2[0] > currentRowSize - 1 || cell2[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            if (cellout[0] > currentRowSize - 1 || cellout[1] > currentColSize - 1)
            {
                Console.WriteLine("One of these cells are not exist on the excelTable");
                Console.WriteLine("");
                return;
            }
            int ans = 0;
            string ansstr = "";
            string val1 = excelTable[cell1[0], cell1[1]];
            string val2 = excelTable[cell2[0], cell2[1]];
            string type1 = dataTypes[cell1[0], cell1[1]];
            string type2 = dataTypes[cell2[0], cell2[1]];

            if (dataTypes[cell1[0], cell1[1]] == "unassigned" || dataTypes[cell2[0], cell2[1]] == "unassigned")
            {
                Console.WriteLine("Illegal operation! One of the input cells is unassigned.");
                return;
            }

            if (type1 == type2)
            {
                if (type1 == "integer")
                {
                    ans = Convert.ToInt32(val1) - Convert.ToInt32(val2);
                    excelTable[cellout[0], cellout[1]] = Convert.ToString(ans);
                    dataTypes[cellout[0], cellout[1]] = "integer";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + Convert.ToString(ans));
                }
                if (type1 == "string")
                {

                    ansstr = removinglettersstring(val1, val2);
                    excelTable[cellout[0], cellout[1]] = ansstr;
                    dataTypes[cellout[0], cellout[1]] = "string";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + ansstr);

                }
            }
            if (type2 != type1)
            {

                if (type1 == "integer" && type2 == "string")
                {
                    int number = Convert.ToInt32(val1);
                    if (number > 126 || number < 33)
                    {
                        Console.WriteLine("Illegal operation!");
                        Console.WriteLine("The integer parameter should be in [33, 126]");
                        return;
                    }

                    char a = (char)(number);
                    ansstr = removinglettersint(val2, a);
                    excelTable[cellout[0], cellout[1]] = ansstr;
                    dataTypes[cellout[0], cellout[1]] = "string";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + ansstr);

                }
                if (type1 == "string" && type2 == "integer")
                {
                    int number = Convert.ToInt32(val2);
                    if (number > 126 || number < 33)
                    {
                        Console.WriteLine("Illegal operation!");
                        Console.WriteLine("The integer parameter should be in [33, 126]");
                        return;
                    }

                    char a = (char)(number);

                    ansstr = removinglettersint(val1, a);


                    excelTable[cellout[0], cellout[1]] = ansstr;
                    dataTypes[cellout[0], cellout[1]] = "string";
                    Console.WriteLine("Operation is done!");
                    Console.WriteLine(c1 + ": " + val1);
                    Console.WriteLine(c2 + ": " + val2);
                    Console.WriteLine("Output: " + ansstr);

                }
            }
        }
        static void DecodeCommand(string input)
        {
            if (input == null)
            {
                Console.WriteLine("Please enter valid input you can not enter null input");
                Console.WriteLine("");
                return;
            }
            if (input.Length < 2)
            {
                Console.WriteLine("Please enter valid input !");
                Console.WriteLine("");
                return;
            }
            int openParanIndex = input.IndexOf('(');

            if (openParanIndex == -1)
            {
                string firstNumIndex = Convert.ToString(input[1]);
                if (input.Length > 2)
                {
                    string secondNumIndex = Convert.ToString(input[2]);
                    if (input.Length > 3 || (!(IntegerCheck(firstNumIndex) || IntegerCheck(secondNumIndex))))
                    {
                        Console.WriteLine("Please enter valid input !");
                        Console.WriteLine("");

                        return;
                    }
                }


                if (input.Length > 3 || (!IntegerCheck(firstNumIndex)))
                {
                    Console.WriteLine("Please enter valid input !");
                    Console.WriteLine("");

                    return;
                }

                string letter = input.Substring(0, 1);
                letter = letter.ToUpper();
                string number = input.Substring(1);
                int rowIndex = Convert.ToInt32(number) - 1;
                int colIndex = (int)letter[0] - 65;
                if (IntegerCheck(Convert.ToString(input[0])))
                {
                    Console.WriteLine("Please enter valid input !");
                    Console.WriteLine("");
                    return;
                }
                if (rowIndex >= 15 || colIndex >= 10 || rowIndex < 0 || colIndex < 0)
                {
                    Console.WriteLine("Out of bounds exception!");
                    Console.WriteLine("Error: Maximum limit is 15 rows and 10 columns (A-J).");
                    Console.WriteLine("");
                    return;
                }
                if (rowIndex > currentRowSize - 1 || colIndex > currentColSize - 1 || colIndex < 0 || rowIndex < 0)
                {
                    Console.WriteLine($"Error: Cell is outside the current table size ({currentRowSize}x{currentColSize}). Use 'AddRow' or 'AddColumn'.");
                    Console.WriteLine("");
                    return;

                }
                Console.WriteLine(excelTable[rowIndex, colIndex]);
                return;
            }
            if (openParanIndex != -1)
            {
                string command = input.Substring(0, openParanIndex);
                command = command.Trim();
                string paramsRaw = input.Substring(openParanIndex + 1).Trim().TrimEnd(')');
                string[] param = paramsRaw.Split(',');

                for (int i = 0; i < param.Length; i++)
                {
                    param[i] = param[i].Trim();
                }
                switch (command)
                {
                    default:
                        Console.WriteLine("Unknown command! Please check your syntax.");
                        break;
                    case "AssignValue":
                        if (param.Length != 3)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 3");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 3)
                        {

                            AssignValue(param[0], param[1], param[2]);
                        }

                        break;

                    case "ClearCell":

                        if (param.Length != 1)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 1");
                            Console.WriteLine("");
                            break;
                        }

                        if (param.Length == 1)
                        {
                            ClearCell(param[0]);
                            Console.WriteLine("Operation is done!");

                        }

                        break;

                    case "ClearAll":

                        ClearAll();
                        Console.WriteLine("Operation is done!");



                        break;

                    case "AddRow":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 2)
                        {
                            AddRow(param[0], param[1]);

                        }

                        break;
                    case "AddColumn":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 2)
                        {
                            AddColumn(param[0], param[1]);

                        }
                        break;
                    case "Copy":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 2)
                        {
                            Copy(param[0], param[1]);
                        }

                        break;
                    case "CopyColumn":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 2)
                        {
                            CopyColumn(param[0], param[1]);
                        }

                        break;
                    case "CopyRow":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 2)
                        {
                            CopyRow(param[0], param[1]);
                        }
                        break;

                    case "X":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 2)
                        {
                            X(param[0], param[1]);

                        }

                        break;
                    case "XColumn":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }

                        if (param.Length == 2)
                        {

                            XColumn(param[0], param[1]);

                        }

                        break;

                    case "XRow":
                        if (param.Length != 2)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 2");
                            Console.WriteLine("");
                            break;
                        }

                        if (param.Length == 2)
                        {

                            XRow(param[0], param[1]);

                        }

                        break;

                    case "*":
                        if (param.Length != 3)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 3");
                            Console.WriteLine("");
                            break;
                        }

                        if (param.Length == 3)
                        {
                            mult(param[0], param[1], param[2]);
                        }

                        break;

                    case "+":
                        if (param.Length != 3 && param.Length != 4)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 3 or 4");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 3)
                        {
                            sum3param(param[0], param[1], param[2]);
                        }
                        else if (param.Length == 4)
                        {
                            sum4param(param[0], param[1], param[2], param[3]);
                        }

                        break;
                    case "/":
                        if (param.Length != 3)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 3");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 3)
                        {
                            divide(param[0], param[1], param[2]);
                        }

                        break;
                    case "-":
                        if (param.Length != 3)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 3");
                            Console.WriteLine("");
                            break;
                        }

                        if (param.Length == 3)
                        {
                            substract(param[0], param[1], param[2]);
                        }

                        break;
                    case "#":
                        if (param.Length != 3)
                        {
                            Console.WriteLine("Parameter number is wrong.It should be 3");
                            Console.WriteLine("");
                            break;
                        }
                        if (param.Length == 3)
                        {
                            hash(param[0], param[1], param[2]);
                        }

                        break;

                }
            }
        }

        static void Main(string[] args)
        {
            InitArrays();

            readdata();
            while (true)
            {


                Console.Write(">> ");
                string cmd = Console.ReadLine();
                Console.WriteLine();
                cmd = cmd.Trim();
                DecodeCommand(cmd);
                PrintTable();
                Console.WriteLine("");

                writedata();

            }

        }
    }
}
