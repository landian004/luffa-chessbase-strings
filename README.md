# luffa-chessbase-strings
翻译和修改ChessBase中文语言的gui软件
<img width="1791" height="1247" alt="image" src="https://github.com/user-attachments/assets/5c275768-72fb-4b66-bc6a-aa6f3b48b59f" />
作者：AI、854982774@qq.com（我） 日期：2026年

软件功能：

①解决ChessBase和Fritz中文界面出现NOTEXT-XXXX文字问题（生成新Chinese按钮一键解决）

②方便搜索、修改翻译。

适用于：CB16，17，18，26， Fritz17、18、19、20、21， 这8个版本官方支持中文。后续更新的CB和Fritz版本理论上也支持，但未测试。

使用教程：https://www.bilibili.com/video/BV1LZvpB1ERS/

Software features:

① Fixes the issue where Chinese UI in ChessBase and Fritz shows "NOTEXT-XXXX" text (one-click fix via a generated new "Chinese" button)

② Makes it easy to search for and edit translations.

The principle behind the "Generate New Chinese" button:

First, it is recommended to replace the ChessBase or Fritz Chinese folder provided in this project with the one in the Messages directory you want to modify, and then run "Generate New Chinese." This ensures that a usable Chinese_new can be generated.

English serves as the reference baseline for generating the new Chinese (do not make any changes to the English folder).

"Generate New Chinese" first scans all file names ending with the .strings suffix in the English folder. Since Openings, Cities, and Mia do not conform to the modification format standards, these three files are excluded first.

For all other .strings files:

① First, check whether the Chinese folder has the corresponding file. If not, copy the corresponding file from English to the Chinese directory. Therefore, apart from the three excluded files, all other files will first be identical at the file level.

② Then, determine whether the uppercase symbols in each file match between Chinese and English.

If they match, the original Chinese file is copied unchanged to the new Chinese (Chinese_new) directory.

If they do not match, NOTEXT-XXXX typically appears in the software's Chinese interface. In this case, the English uppercase symbols are used as the standard (since English never has issues), and all Chinese translations in the file content are written into the newly generated file (e.g., app.strings) according to the uppercase symbol parts that match the English uppercase symbols.

For cases where Chinese and English differ, either the Chinese version is missing the uppercase symbol, or the Chinese uppercase symbol spelling does not match. In all such cases, the English original text is used and written into the corresponding .strings file in the new Chinese folder.

Now that we understand this principle, we can safely use the "Generate New Chinese" button — it will not lose our previous translation work. Just be careful not to click "Generate New Chinese" again, as it would overwrite our modified translations (all translation results are saved in the Chinese_new folder).

-------
“生成新Chinese”按钮的原理：

首先，建议把本项目中提供的ChessBase或Fritz的Chinese文件夹先行替换到要进行修改的Messages目录中，然后再运行“生成新Chinese”。这样确保能生成可用的Chinese_new。

English是作为生成新Chinese的对照基准(English文件夹不要做任何改动)。

“生成新Chinese” 首先扫描English文件夹内的所有以.strings为后缀的文件名字，因为Openings、Cities和Mia不符合修改的格式标准，所以首先把这3个排除出去。

对其它的所有strings文件：

① 首先判断Chinese文件夹是否有相应文件，如果没有就把English中的对应文件先复制到Chinese目录下，所以，除了排除出去的3个文件外，其它的都会文件层面上首先一致。

② 然后再逐个判断每个文件的大写符号是否相同，

如果相同，则原Chinese的该文件原封不动复制到新Chinese（Chinese_new）目录去。

如果不同，则通常在软件中文界面下就会出现NOTEXT-XXXX，这时会以英文的大写符号为准（因为英文的永远不会有问题），把所有文件内容的中文的翻译按英文大写符号相同的大写符号部分一一写到新生成的该文件比如app.strings文件内容中，

对于中、英文不同的情况，要么是中文中缺少这个大写符号、要么是中文的大写符号拼写不符了，这些不符的情况都会取英文的原文写入到新Chinese中的相应strings文件内容中。

所以知道这个原理后，我们就可以放心地使用“生成新Chinese”按钮，它不会丢失我们之前翻译的成果。注意不要再次点击“生成新Chinese”，否则会覆盖我们修改的翻译成果（翻译成果都是保存在Chinese_new文件夹内的）。

