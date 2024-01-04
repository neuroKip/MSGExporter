using static System.Net.Mime.MediaTypeNames;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Text;
using System.Security.Cryptography;
using System.Diagnostics;
using System;
using System.Linq;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Drawing;

namespace MsgExporter
{
    internal class Program
    {
        /*** DEBUG ***/
        static string curMsgBeingParsed = ""; // Keep track of currently open .MSG for debug purposes
        // Some constants I found by going through all .MSG files
        const int maxEntrySizeKnown = 505; // Biggest know entry size (in bytes) in a .MSG file
        const int maxCharPerLineKnown = 35; // NOTE: This might not be exact due to how I extracted it (see debugCharPerLineCount)
        const int maxNumOfEntriesKnown = 594; // Biggest number of entries find in a .MSG file

        // This will be filled up with the FONT
        static Dictionary<UInt16, char> kanji = new Dictionary<UInt16, char>();
        static Dictionary<UInt16, string> specialCharacters = new Dictionary<UInt16, string>();

        enum Mode
        {
            None,

            ExtractFromMsg, // Given a .MSG, extract its text to UTF
            ConvertToMsg    // Convert a plain text file to a .MSG
        }
        static Mode currentMode = Mode.None;

        /*** EXPORT FLAGS ***/
        // When extracting text from a .MSG we can set extra flags
        static bool exportTSV = false; // Save output as Tab Separated Values, handy to paste/import into Google Sheets

        static bool printHeaderInfo = false; // This will print info about the .MSG header, useful for DEBUG
        static bool printDelimitersText = false; // Print begin and end position of each text entry in the .MSG
        static bool printHexFromMsg = false; // Print char hex value from .MSG
        static bool printHexFromFont = false; // Print char hex value from Font file
        static bool printEndOfLine = false; // Print end of line too {EEA3BF}

        static bool skipUserInput = false; // Skip waiiting for key, handy to pipe this tool

        static string fontPath = ""; // REQUIRED Path to the font table extracted from font file
        static string inputPath = ""; // REQUIRED This is the entry file, can be a directory or a single .MSG/.TSV depending on the Mode
        static string outDirPath = ""; // OPTIONAL ./Exported/ or ./Converted/ if out dir it not user defined, exported/converted files are named same as input files

        const string helpText = "\t-export\t Extract text from a .msg to .txt" +
            "\t-export -font path/to/font/file -input /path/to/.msg" +
            "\t-convert\t Convert a .tsv containing text to a .msg\n\t\tTSV Format should be single column \"String1\"\\n\"String2\"" +
            "\t-convert -font path/to/font/file -input /path/to/.tsv" +
            "\t-input path\tif -export a .msg is expected, if -convert a .tsv instead, directories are also allowed" +
            "\t-font path\tPath to a .tsv containing chars extracted from the font file\nFormat is:\nhex\tchar\n4D\tM\n482F\tR" +
            "\t-out path\tUser defined out path, otherwise ./Exporter/ or ./Converted/" +
            "\t-tsv\tExport to a .tsv instead of .txt" +
            "\t-debug\tPrint extra debug info to the -export .txt" +
            "\t-skip Do not wait for an input key at the end";

        static int Main(string[] args)
        {
            if (ProcessArguments(args) == false)
            {
                Console.WriteLine("FAILED");
                WaitForUserInput(); // Error encountered, do not close cmd right away

                return -1;
            }

            // Create dir where to export stuff
            DirectoryInfo outDir = new DirectoryInfo(outDirPath);
            if (outDir.Exists == false)
            {
                outDir.Create();
            }

            /* READ FONT FROM FILE */
            List<string> fontDuplicateEntries = new List<string>();
            if (CreateFontFromFile(fontPath, out fontDuplicateEntries))
            {
                Console.WriteLine($"FONT CharSet parsed, {kanji.Count} entries found, {specialCharacters.Count} special charactes found");
            }
            else
            {
                // All entries in the FONT.tsv should be unique, if not, abort export and create a txt with a list of duplicates
                string duplicatesOutFile = Path.Combine(outDir.FullName, "DUPLICATES.txt");
                Console.WriteLine($"ERROR: Found duplicate indices in the FONT file, see {duplicatesOutFile} for info.");

                using (StreamWriter outFileStream = new StreamWriter(duplicatesOutFile))
                {
                    outFileStream.WriteLine("DUPLICATES");
                    outFileStream.Write(string.Join("\n", fontDuplicateEntries));
                }

                WaitForUserInput();

                return -2;
            }

            // Find files in the given input path
            string inExt = currentMode == Mode.ExtractFromMsg ? ".msg" : ".tsv";
            FileInfo[] inputFiles = FindFiles(inputPath, inExt);
            if (inputFiles.Length == 0)
            {
                Console.WriteLine($"No files with {inExt} extention found in given -input path!");
                WaitForUserInput();

                return -3;
            }

            /* CONVERT GIVEN .TSV FILE TO .MSG */
            if (currentMode == Mode.ConvertToMsg)
            {
                foreach (FileInfo tsvFile in inputFiles)
                {
                    string outFileName = tsvFile.Name.Replace(tsvFile.Extension, ".msg");
                    FileInfo outFile = new FileInfo(Path.Combine(outDir.FullName, outFileName));
                    outFile.Delete();
                    FileStream msgFileStream = outFile.Create();

                    Console.WriteLine($"Converting {tsvFile.Name}...");

                    ConvertToMsg(inputPath, ref msgFileStream);

                    Console.WriteLine($"Converted to {outFile}\n");

                    msgFileStream.Close();
                    msgFileStream.Dispose();
                }
            }
            /* EXPORT GIVEN .MSG FILE TO PLAIN TEXT */
            else if (currentMode == Mode.ExtractFromMsg)
            {
                string outExt = exportTSV ? ".tsv" : ".txt";

                foreach (FileInfo msgFile in inputFiles)
                {
                    string outFileName = msgFile.Name.Replace(msgFile.Extension, outExt);
                    string outFile = Path.Combine(outDir.FullName, outFileName);

                    using (StreamWriter outFileStream = new StreamWriter(outFile))
                    {
                        curMsgBeingParsed = msgFile.Name;
                        Console.WriteLine($"Extracting {curMsgBeingParsed}...");

                        int numOfEntries;
                        List<string> lines = ExtractMsg(msgFile.FullName, out numOfEntries);
                        outFileStream.Write(string.Join("\n", lines));

                        Console.WriteLine($"{numOfEntries + 1} Text Entries found");
                        Console.WriteLine($"Extracted to {outFile}\n");
                    }
                }
            }

            Console.WriteLine("DONE");
            WaitForUserInput();

            return 0;
        }

        static bool ProcessArguments(string[] args)
        {
            if (args.Length == 0)
            {
                // Print help
                Console.WriteLine(helpText);
                return false;
            }

            // Parse cmd arguments
            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i].ToLower();

                switch (arg)
                {
                    case "-h":
                    case "-help":
                        {
                            Console.WriteLine(helpText);
                            return false;
                        }

                    case "-input":
                        {
                            inputPath = args[i + 1];
                        }
                        break;

                    case "-out":
                        {
                            outDirPath = args[i + 1];
                        }
                        break;

                    case "-font":
                        {
                            fontPath = args[i + 1];
                        }
                        break;

                    case "-export":
                        {
                            currentMode = Mode.ExtractFromMsg;
                        }
                        break;

                    case "-convert":
                        {
                            currentMode = Mode.ConvertToMsg;
                        }
                        break;

                    case "-tsv":
                        {
                            exportTSV = true;
                        }
                        break;

                    case "-debug":
                        {
                            printHeaderInfo = true; // This will print info about the .MSG header, useful for DEBUG
                            printDelimitersText = true; // Print begin and end position of each text entry in the .MSG
                            printHexFromMsg = true; // Print char hex value from .MSG
                            printHexFromFont = true; // Print char hex value from Font file
                            printEndOfLine = true; // Print end of line too {EEA3BF}
                        }
                        break;

                    case "-skip":
                        {
                            skipUserInput = true;
                        }
                        break;
                }
            }

            // Sanitize args parsing results

            if (fontPath == "" || File.Exists(fontPath) == false)
            {
                // We can not do anything if we do not have a FONT file!
                Console.WriteLine("-font required, see README for font file format");
                return false;
            }

            if (currentMode == Mode.None)
            {
                // We need to define which mode we want to use!
                Console.WriteLine("No mode selected, -export or -convert required");
                return false;
            }

            if (outDirPath == "")
            {
                outDirPath = currentMode == Mode.ExtractFromMsg ? "./Exported/" : "./Converted/";
            }

            return true;
        }

        static void WaitForUserInput()
        {
            if (skipUserInput)
                return;

            Console.ReadLine();
        }

        static bool CreateFontFromFile(string fontFilePath, out List<string> fontDuplicateEntries)
        {
            // We expect a '\t' separated TSV (that's the format that Google Sheets stores to clipboard)
            // First column should be FONT index code, second is the actual character
            // E.G. FONT.tsv contains:
            //21	れ
            //22	+
            //23	0
            //24	{24}
            //25	も
            //8224	私

            fontDuplicateEntries = new List<string>();

            kanji.Clear();
            specialCharacters.Clear();

            using (StreamReader sr = new StreamReader(fontFilePath))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] split = line.Split("\t");

                    UInt16 charSetIndex = Convert.ToUInt16($"0x{split[0]}", 16);

                    if (split[1].Length == 1)
                    {
                        bool success = kanji.TryAdd(charSetIndex, split[1].First());
                        if (success == false)
                        {
                            fontDuplicateEntries.Add(line);
                        }
                    }
                    else
                    {
                        bool success = specialCharacters.TryAdd(charSetIndex, split[1]);
                        if (success == false)
                        {
                            fontDuplicateEntries.Add(line);
                        }
                    }
                }
            }
            // All entries in the FONT.tsv should be unique, if not, abort export and create a txt with a list of duplicates
            return fontDuplicateEntries.Count() == 0;
        }

        static FileInfo[] FindFiles(string inputPath, string fileExt)
        {
            // Check if the given path is a file or a directory
            if (File.Exists(inputPath))
            {
                // TODO Do some safety checks on the file path
                return new FileInfo[] { new FileInfo(inputPath) };
            }

            if (Directory.Exists(inputPath))
            {
                // Path is a directory, find all files with the extension we're interested in (.msg or .tsv)
                DirectoryInfo filesDir = new DirectoryInfo(inputPath);
                FileInfo[] files = filesDir.GetFiles($"*{fileExt}", SearchOption.AllDirectories);

                if (files.Length == 0)
                {
                    Console.Write($"no {fileExt} files found in specified path!");
                    Console.ReadLine();

                    return files;
                }

                return files;
            }

            WaitForUserInput();

            return new FileInfo[0];
        }

        static List<string> ExtractMsg(string path, out int numOfEntries)
        {
            FileStream msgFileStream = new FileStream(path, FileMode.Open);
            List<string> lines = new List<string>(); // Strings that will go in the .txt file
            numOfEntries = -1; // Counter for parsed entries

            if (exportTSV == false)
            {
                lines.Add("EXPORT FROM " + path.Substring(path.LastIndexOf('\\') + 1) + "\n");
            }

            /****** PARSE HEADER ******/

            byte[] headerByteBuffer = new byte[16];
            int r = msgFileStream.Read(headerByteBuffer, 0, 16);

            int bottomScreenEntry;
            int numOfDelimiters;
            int fileSize;
            bool result = ParseHeader(headerByteBuffer, out bottomScreenEntry, out numOfDelimiters, out fileSize);
            if (result == false)
            {
                msgFileStream.Close();
                return lines;
            }

            if (exportTSV == false && printHeaderInfo)
            {
                lines.Add("HEADER");
                lines.Add($"bottom screen entry {bottomScreenEntry} expected entries {numOfDelimiters} fileSize {fileSize} bytes\n");
            }

            // Grab first text entry delimiters
            byte[] entryDelimitersBuffer = new byte[8];
            r = msgFileStream.Read(entryDelimitersBuffer, 0, 8);
            int firstEntryStart = BitConverter.ToInt32(entryDelimitersBuffer, 0);

            /****** PARSE ENTRIES ******/

            do // Loop through entry delimiters, until the pointer hits the address of the first entry
            {
                // Store entry offets
                numOfEntries += 1;

                // Parse text entry delimiters and size
                int entryOffsetStart = BitConverter.ToInt32(entryDelimitersBuffer, 0);
                int entryOffsetEnd = BitConverter.ToInt32(entryDelimitersBuffer, 4);
                int entrySize = entryOffsetEnd - entryOffsetStart;

                if (exportTSV == false)
                {
                    lines.Add($"***ENTRY {numOfEntries + 1}***");

                    if (printDelimitersText)
                    {
                        //Console.WriteLine( lines.Last() );
                        lines.Add($"START {entryOffsetStart.ToString("X2")} END {entryOffsetEnd.ToString("X2")} SIZE {entrySize}");
                        lines.Add($"RAW DELIMITERS {BitConverter.ToString(entryDelimitersBuffer)}");
                        //Console.WriteLine( lines.Last() );
                    }
                }

                // Move stream pointer to beginning of entry
                long prevPos = msgFileStream.Position;
                msgFileStream.Seek(entryOffsetStart, SeekOrigin.Begin);

                /*** Read entry and cache it ***/

                byte[] stringBuffer = new byte[entrySize];
                r = msgFileStream.Read(stringBuffer, 0, entrySize);
                string rawTextEntry = BitConverter.ToString(stringBuffer);

                // Go through the byte buffer and convert to plain text
                List<Tuple<byte[], int>> convertedIndices;
                string stringEntry = ParseEntry(stringBuffer, out convertedIndices);

                // Print plain text entry
                if (exportTSV)
                {
                    lines.Add($"\"{stringEntry}\"");
                }
                else
                {
                    lines.Add($"\nTEXT\n{stringEntry}");
                }

                // Print also the byte buffer for debug purposes
                if (exportTSV == false && printHexFromMsg)
                {
                    lines.Add("\nHEX FROM MSG\n" + rawTextEntry);
                }

                // Print converted indices to simply searches in FONT
                if (printHexFromFont)
                {
                    foreach (var conversion in convertedIndices)
                    {
                        string twoBytesCode = BitConverter.ToString(conversion.Item1);
                        string fontIndex = conversion.Item2.ToString("X2");
                        rawTextEntry = rawTextEntry.Replace(twoBytesCode, fontIndex);
                    }
                    //string eol = printEndOfLine ? "{EEA3BF}\n" : "\n";
                    rawTextEntry = rawTextEntry.Replace("-EE-A3-BF-", "{EEA3BF}\n");
                    rawTextEntry = rawTextEntry.Replace("EE-A0-81-00", "{EEA08100}");

                    if (exportTSV)
                    {
                        lines[lines.Count - 1] = $"\"{rawTextEntry}\",{lines[lines.Count - 1]}";
                    }
                    else
                    {
                        lines.Add($"\nHEX FROM FONT\n{rawTextEntry}\n\n\n");
                    }
                }

                // Move stream pointer back to the entry delimiters
                msgFileStream.Seek(prevPos, SeekOrigin.Begin);
                r = msgFileStream.Read(entryDelimitersBuffer, 0, 8);

            } while (msgFileStream.Position <= firstEntryStart); // TODO tidy up this loop to use numOfEntries instead

            // Check that the number of entries parsed matches the count from the header
            Debug.Assert((numOfEntries + 1) == numOfDelimiters, $"WRONG AMOUNT OF ENTRIED EXPECTED {numOfDelimiters} FOUND {numOfEntries}");

            msgFileStream.Close();

            return lines;
        }

        static bool ParseHeader(byte[] headerByteBuffer, out int bottomScreenEntry, out int numOfDelimiters, out int fileSize)
        {
            //Header format is:
            //00 00 00 00 54 45 58 54 02 1f 00 00 2a 07 00 00
            //00 00 00 00 (4 zeroed out bytes)
            //54 45 58 54 (TEXT, constant to recognize .MSG files)
            //02 (text entry that bottom screen starts with?)
            //1f 00 (number of text entries)
            //00 (1 zeroed byte?)
            //2a 07 (file size in bytes)
            //00 00 (2 zeroed bytes?)

            // TEXT
            // We start at 4 since the hader starts with 0000
            if (System.Text.Encoding.UTF8.GetString(headerByteBuffer, 4, 4) != "TEXT")
            {
                Console.WriteLine("Failed to read header: TEXT not found!");
                bottomScreenEntry = numOfDelimiters = fileSize = -1;

                return false;
            }

            // text entry that bottom screen starts with?
            bottomScreenEntry = BitConverter.ToUInt16(headerByteBuffer, 8);
            // Second byte should be the number of entries
            numOfDelimiters = BitConverter.ToUInt16(headerByteBuffer, 9);

            //maxNumOfEntriesKnown = Math.Max( maxNumOfEntriesKnown, numOfDelimiters );

            // File size
            fileSize = BitConverter.ToUInt16(headerByteBuffer, 12);

            return true;
        }

        static string ParseEntry(byte[] stringBuffer, out List<Tuple<byte[], int>> convertedIndices)
        {
            string stringEntry = "";
            convertedIndices = new List<Tuple<byte[], int>>();
            int debugCharPerLineCount = 0;

            for (int i = 0; i < stringBuffer.Length; i++)
            {
                if (stringBuffer[i] == 0x00)
                    continue;

                if (stringBuffer[i] < 0xC0) // We assume that anything below 0xC0 is not a kanji, AKA one byte char
                {
                    UInt16 kanjiIndex = Convert.ToUInt16(stringBuffer[i]);
                    char singleByteChar;
                    bool charFound = kanji.TryGetValue(kanjiIndex, out singleByteChar);

                    if (charFound)
                    {
                        debugCharPerLineCount += 1;

                        stringEntry += singleByteChar;
                    }
                    else if (specialCharacters.ContainsKey(kanjiIndex)) // We haven't found this character, look it up in the special list
                    {
                        stringEntry += specialCharacters[kanjiIndex];
                    }
                    else // We do not know this character, print it raw
                    {
                        stringEntry += $"{{{stringBuffer[i].ToString("X2")}}}";
                    }
                }
                else if (stringBuffer[i] == 0xEE) // Handle opcode, ignore bytes after it
                {
                    // First two bytes of opcode can be either 0xEEA3 or 0xEEA0

                    if (stringBuffer[i + 1] == 0xA3)
                    {
                        if (stringBuffer[i + 2] == 0xBF) // End of Line code is EE-A3-BF
                        {
                            //debugCharPerLineCount += 1;
                            //maxCharPerLineKnown = Math.Max( maxCharPerLineKnown, debugCharPerLineCount );
                            debugCharPerLineCount = 0;

                            if (printEndOfLine)
                            {
                                stringEntry += "{" + BitConverter.ToString(stringBuffer, i, 3).Replace("-", "") + "}";
                            }
                            stringEntry += "\n";
                        }
                        else
                        {
                            stringEntry += "{" + BitConverter.ToString(stringBuffer, i, 3).Replace("-", "") + "}";
                        }

                        i += 2;
                    }
                    else if (stringBuffer[i + 1] == 0xA0) // Player character name code EE-A0-81-00
                    {
                        stringEntry += "{" + BitConverter.ToString(stringBuffer, i, 4).Replace("-", "") + "}";

                        i += 3;
                    }
                    else // Unknown opcode
                    {
                        Debug.Assert(false, "WE DO NOT KNOW THIS OPCODE! " + BitConverter.ToString(stringBuffer, i, 4));
                    }
                }
                else // this is a kanji, calculate index in kanji list (FONT)
                {
                    byte[] kanjiBytes = new byte[2]
                    {
                        stringBuffer[i],
                        stringBuffer[i + 1],
                    };

                    // Use bitmask to convert from .msg byte code to FONT byte code (set 7th and 14th to 0)
                    kanjiBytes[0] &= 0xBF; //1011 1111
                    kanjiBytes[1] &= 0x7F; //0111 1111
                    kanjiBytes = kanjiBytes.Reverse().ToArray();

                    UInt16 kanjiIndex = BitConverter.ToUInt16(kanjiBytes);
                    //Tuple<UInt16, string> charEntry = kanji.Find( k => k.Item1 == kanjiIndex );
                    char twoByteChar;
                    bool kanjiFound = kanji.TryGetValue(kanjiIndex, out twoByteChar);
                    if (kanjiFound)
                    {
                        stringEntry += twoByteChar;

                        debugCharPerLineCount += 1;
                    }
                    else if (specialCharacters.ContainsKey(kanjiIndex)) // We haven't found this character, look it up in the special list
                    {
                        stringEntry += specialCharacters[kanjiIndex];
                    }
                    else // We do not know this character, print it raw
                    {
                        stringEntry += "{" + kanjiIndex.ToString("X2") + "}";
                    }

                    convertedIndices.Add(new Tuple<byte[], int>(new byte[] { stringBuffer[i], stringBuffer[i + 1] }, kanjiIndex));

                    i += 1; // Kanji are two bytes so skip next one
                }
            }

            return stringEntry;
        }

        static void ConvertToMsg(string translatedTSVPath, ref FileStream outMsgFileStream)
        {
            // Parse translated entries from .TSV,
            // we expect a file with one column of C style strings separated by '\n' 
            //(E.G. "entry1"\n"entry2"...)
            //"{EEA08100}:\n
            //Oh, I know.\n
            //Point check.\n
            //Go first."

            List<string> romanjiEntries = new List<string>();

            // Parse .TSV, split string on '"' and replace '\n' with {EEA3BF}
            using (StreamReader sr = new StreamReader(translatedTSVPath))
            {
                // Text Entries can have more than one line, so we need to concatened them until we hit the end of string '"'
                string currentTextEntry = ""; // current Entry we are working on
                string line; // current line we are on

                while ((line = sr.ReadLine()) != null)
                {
                    // TODO Right now we skip empty lines, but there are .msg files that might have actual empty entries
                    // (maybe used for pausing text animations in certain spots?), this should be investigated at some point
                    if (line == "")
                        continue;

                    // ReadLine() splits on '\n', but we still need it for the export, so replace it with .msg bytecode
                    currentTextEntry += line + "{EEA3BF}";
                    if (line.Last() == '"') // Check if we are at the end of an entry
                    {
                        // Add to Entries and clear entry cache
                        romanjiEntries.Add(currentTextEntry.Replace("\"", "")); // Remove quotations (should be first and last string char)
                        currentTextEntry = "";
                    }
                }
            }

            string debugStr = "";

            /*** WRITE .MSG HEADER ***/

            UInt16 numOfEntries = (UInt16)romanjiEntries.Count();
            byte[] numEntriesByteBuffer = BitConverter.GetBytes(numOfEntries);
            // Add a dummy header for now
            byte[] headerByteBuffer = new byte[16]
            {
                0x00, 0x00, 0x00, 0x00,
                0x54, 0x45, 0x58, 0x54, // TEXT
                0x02, // TODO No clue what to put in here yet when importing translated text!
                numEntriesByteBuffer[0], numEntriesByteBuffer[1], // number of entries (2 bytes)
                0x00, // 1 byte padding?
                0x00, 0x00, // File size (2 bytes)
                0x00, 0x00, // 2 byte padding?
            };
            outMsgFileStream.Write(headerByteBuffer);



            /*** TEXT ENTRIES DELIMITERS ***/

            long currentDelimiterPos = outMsgFileStream.Position;

            // Pre allocate Text Entry Delimiters
            // Entry Delimiters format (8 bytes total):
            // (E.G. BD 00 00 00 5D 01 00 00)
            // BD 00 00 00 (2 bytes Entry start)(2 bytes buffer?)
            // 5D 01 00 00 (2 bytes Entry end)(2 bytes buffer?)
            int delimiterSize = 8;
            byte[] entryDelimitersBuffer = new byte[delimiterSize * romanjiEntries.Count()];
            outMsgFileStream.Write(entryDelimitersBuffer);

            int entryDelimitersBufferSize = entryDelimitersBuffer.Length * delimiterSize;



            /*** WRITE ENTRIES ***/

            // Let's flip the dictionaries so we can get FONT indices by char/string
            List<KeyValuePair<UInt16, char>> fontChars = kanji.ToList();
            List<KeyValuePair<UInt16, string>> specialFontChars = specialCharacters.ToList();

            // Write entries to out stream
            for (int j = 0; j < romanjiEntries.Count; j++)
            {
                // Start position of the Text Entry, we'll write it out to the delimiters list after writing the entry
                long entryStart = outMsgFileStream.Position;
                string roamnjiEntry = romanjiEntries[j];

                // Go through each charater in the entry, and convert from FONT indices to .msg byte codes
                for (int i = 0; i < roamnjiEntry.Length; i++)
                {
                    char currentChar = roamnjiEntry[i];
                    UInt16 fontIndex;

                    // Check if we hit a bytecode
                    if (currentChar == '{')
                    {
                        // Grab successive characters in the Text Entry till we hit '}' (E.G. {EEA3BF})
                        int specialEnd = roamnjiEntry.IndexOf('}', i) - 1;
                        string bytecode = roamnjiEntry.Substring(i + 1, specialEnd - i);

                        try
                        {
                            // Pick two char at the time and try converting to single byte (E.G. first one would be EE)
                            for (int b = 0; b < bytecode.Length; b += 2)
                            {
                                //UInt16 ttt = Convert.ToUInt16( "EEA3BF", 16 );
                                string bbb = bytecode.Substring(b, 2);

                                // If we hit a wrong substring Convert.ToUInt16 will throw an exception
                                UInt16 aaa = Convert.ToUInt16(bbb, 16);
                                byte[] singleByte = BitConverter.GetBytes(aaa);
                                outMsgFileStream.WriteByte(singleByte[0]);
                            }

                            fontIndex = 0;
                            i += bytecode.Length + 1; // + 1 to skip '}'

                            continue;
                        }
                        catch (Exception)
                        {
                            // Can't convert to byte, treat as normal char (E.G. situations like {|} from MsgLabels_00system.msg could cause this)
                            KeyValuePair<UInt16, char> charEntry = fontChars.Find(pair => pair.Value == currentChar);
                            fontIndex = charEntry.Key;
                        }
                    }
                    else
                    {
                        // Grab the FONT index by character
                        KeyValuePair<UInt16, char> charEntry = fontChars.Find(pair => pair.Value == currentChar);
                        // TODO check if we found nothing!
                        fontIndex = charEntry.Key;
                    }

                    // Convert index to bytes
                    byte[] fontByteCode = BitConverter.GetBytes(fontIndex);
                    fontByteCode = fontByteCode.Reverse().ToArray();

                    // Check if we hit a one or two bytes char
                    if (fontByteCode[0] == 0x00)
                    {
                        outMsgFileStream.WriteByte(fontByteCode[1]);

                        debugStr += fontByteCode[1].ToString("X2") + "-";
                    }
                    else
                    {
                        // Convert FONT two bytes index to .msg bytecode
                        fontByteCode[0] |= 0x40; //0100 0000
                        fontByteCode[1] |= 0x80; //1000 0000

                        outMsgFileStream.Write(fontByteCode, 0, 2);

                        debugStr += BitConverter.ToString(fontByteCode).Replace("-", "") + "-";
                    }
                }

                // We're done writing the Text Entry,
                // now move back to the top of the file and write Entry start and end positions in the file

                // Termiante entry with 00 00
                outMsgFileStream.Write(new byte[2] { 0x00, 0x00 });

                long entryEnd = outMsgFileStream.Position;

                // Move stream pointer back to entry delimiters list
                outMsgFileStream.Seek(currentDelimiterPos, SeekOrigin.Begin);

                // Write at which byte the entry we just wrote starts at
                byte[] startBytes = BitConverter.GetBytes((int)entryStart);
                outMsgFileStream.Write(startBytes);

                // Save the end of the entry (equivalent to start + size in bytes)
                byte[] endBytes = BitConverter.GetBytes((int)entryEnd);
                outMsgFileStream.Write(endBytes);

                // Cache position of the next Entry Delimeters and move back at the end of current Text Entry
                currentDelimiterPos = outMsgFileStream.Position;
                outMsgFileStream.Seek(entryEnd, SeekOrigin.Begin);
            }

            // Write file size to header
            long eof = outMsgFileStream.Position;
            outMsgFileStream.Seek(12, SeekOrigin.Begin);
            byte[] eofBytes = BitConverter.GetBytes((int)eof);
            outMsgFileStream.Write(eofBytes);
        }
    }
}
