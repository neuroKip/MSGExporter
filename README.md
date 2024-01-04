-----------------------------------------------------------------------
ABOUT (V 1.1)
-----------------------------------------------------------------------
This tool is used to export text from a .msg binary file, a text format used in the NDS game Front Mission 2089: Border of Madness (FM2089).
In addition the tool is also capable of converting plain text strings to a .msg.

Note that this file format is used by other NDS games, but this tool was created specificaly for the above mentioned game (See REAMRKS section for more info).

github:
-  https://github.com/neuroKip/MSGExporter

Front Mission 2089 Border of Madness Translation Project Discord:
-  [https://discord.gg/8ZD67Xxt](https://discord.gg/2ehkcvSuuk)

-----------------------------------------------------------------------
REQUIREMENTS
-----------------------------------------------------------------------
* .NET 6.0 Runtime is required to run the tool.
* A Tab Separated Values file (.tsv) listing the characters used in the in game font (see FONT TABLE FILE section for more info)


-----------------------------------------------------------------------
HOW TO USE THE TOOL
-----------------------------------------------------------------------
Run the .exe with `-help` for a more detailed list of commands.

---------------------------
HOW TO EXPORT
---------------------------
Extract a .msg to plain text:
`-export -font "path/to/font/table.tsv" -input "path/to/file.msg"`

`-export` This will tell the tool to extract text from a .msg to a .txt

`-font` The tool needs a look up table to be able to extract the text (see FONT TABLE FILE section for info on how to create one)

`-input` Can be either a path to a .msg, or a directory containing multiple files.

By default -export will create a .txt file, if you want something easier to process or paste into a spreadsheet, the tool can also export to a Tab Separated Values file (.tsv) by using the argument -tsv.
Note that even though the file format is .tsv, the tool just exports the text in a single column separated by \n.

---------------------------
HOW TO CONVERT
---------------------------
Create a .msg from plain text:
`-convert -font "path/to/font/table.tsv" -input "path/to/plain/text.tsv"`

`-convert` Tells the tool to convert a Tab Separated Values file (.tsv) to a .msg
`-font` The tool needs a look up table to be able to convert from plain text to hex char values used in the .msg format (see FONT FILE sectiong for info on how to create one)
`-input` Should point to a single .tsv or a directory containing multiple files.

The .tsv format is expected to be the following:
`"first text entry"\n"second text entry with\nmultiple lines"` 

Note that each text entry is wrapped by double quotes "" and are separated by a \n.


-----------------------------------------------------------------------
THE MSG FILE FORMAT
-----------------------------------------------------------------------

An .msg file is split in three parts, Header, Entry Delimiters, Text Entries.
(Special thanks to SCVgeo from gbatemp.net forums for figuring out the delimiters)

---------------------------
HEADER
---------------------------
The header size of a .msg is always 16 byte in size:
`00 00 00 00 54 45 58 54 02 1f 00 00 2a 07 00 00`

```
00 00 00 00 (4 bytes padding)
54 45 58 54 (TEXT in ascii, constant used to recognize .MSG files)
02 (1 byte, currently uknown)
1f 00 (2 bytes number of text entries in this file)
00 (1 zeroed byte)
2a 07 (2 bytes .msg file size in bytes)
00 00 (2 bytes padding)
```
---------------------------
ENTRY DELIMITERS
---------------------------
Pair of integers delimiting the beginning and end of each text entry relative to the beginning of the file.
Entry delimiters are 8 byte size:
`BD 00 00 00 5D 01 00 00`

```
BD 00 (start of text entry)
00 00 (2 bytes padding)
5D 01 (end of text entry)
00 00 (2 bytes padding)
```
---------------------------
TEXT ENTRIES
---------------------------
The bottom of the .msg file contains a sequence of text entries, most of the time separated by 2 bytes padding. 
A single text entry is composed by one or more lines of characters separated by a next line character (in FM2089 the next line character is always `EE A3 BF`, see REMARKS section for more info). 
A single character can be either 1 or 2 bytes long, in the specific case of FM2089, 2 bytes characters start from C0 00.
The encoding of each character is relative to a font file specific to the game (for example FM2089 font resides in a .fnt file).
While 1 byte characters map directly to the .fnt file, the 2 byte characters need to be converted using a bitmask (Special thanks to andi_ke for figuring out the bitmask).


-----------------------------------------------------------------------
FONT TABLE FILE
-----------------------------------------------------------------------

FM2089 uses a custom character encoding (stored inside a .fnt file), thus when extracting the text from .msg to plain text (and viceversa), we need a look up table of which UTF-8 character corresponds to the same character in the .fnt file.
For instance, while the letter C corresponds to a hex value of 0x43 in UTF-8, in FM2089 its hex value is instead 0x9228.

---------------------------
FONT TABLE FILE FORMAT
---------------------------
A font table file is a Tab Separated Values file (.tsv) containing the list of all characters used by the game, and their corresponding hex value from the game .fnt file.
Format must be one UTF character per line, preceded by their hex value and separated by a \t:
```
21	れ // a 1 byte character
1A	T // another 1 byte character
2BAF	{2BAF} // "special character/icons" character
C1D2	私 // a 2 byte character
```
In the FM2089 .fnt file there are instances of special characters which do not have an equivalent in UTF (such as button icons).
By wrapping their hex value within {} we tell the tool to "pass through" such values when converting from plain text to a .msg.

---------------------------
HOW TO CREATE A FONT TABLE
---------------------------
To create your own font table file, you must extract each character and its hex value from the game's .fnt file.

Luckly we can peek into a .fnt file by using the fantastic tool created by HoRRoR (consolgames.ru) and Djinn (magicteam.net), the Square Enix Remakes Font Editor (SERFontEditor):

- https://www.romhacking.net/utilities/624/

SERFontEditor is built to read/edit/save a .fnt file, but what we are actually interested in, is the hex value of each character (shown at the bottom of the tool, after the W: P:).
The next step is to MANUALLY create the font table file .tsv, and match each character you see in the in the SERFontEditor UI, with one from UTF.
Unfortunately this is a very time consuming process, since each character must be looked up manually in UTF.
A surprisingly effective approach is to take a screenshots of the font, and use an image to text program to extract it, then go through the result and check which characters are incorrect (Special thanks to TJ6 for coming up with this method!).


-----------------------------------------------------------------------
REMARKS
-----------------------------------------------------------------------
While the .msg file format is used in other games, this tool has currently a few limitations that are likely to prevent it from working on .msg files from other games than FM2089 (as of V1.0 of this tool at least).
One of such limitations is the new line character, FM2089 uses 3 bytes to signify when a text entry should go to the next line `EE A3 BF`.
Currently, the tool replaces each instance of `EE A3 BF` found in a .msg with a \n (and viceversa when using `-convert`), which means it will fail if a game uses a different code than the expected one.
