module.exports = {
	ConvertToBijoy
}

/******************************************
	Array containing ASCII to Unicode map 
	for bijoy
*******************************************/

var bijoy_string_conversion_map = {
    "\u09B0\u200C\u09CD\u09AF": "i¨",
    "্র্য": "ª¨",
    "ম্প্র": "¤cÖ",
    "ক্ষ্ম": "²",
    "ক্ক": "°",
    "ক্ট": "±",
    "ক্ত": "³",
    "ক্ব": "K¡",
    "স্ক্র": "¯Œ",
    "ক্র": "µ",
    "ক্ল": "K¬",
    "ক্ষ": "¶",
    
    "ক্স": "·",
    "গু": "¸",
    "গ্ধ": "»",
    "গ্ন": "Mœ",
    "গ্ম": "M¥",
    "গ্ল": "M­",
    "গ্ল": "Mø",
    "ঙ্ক": "¼",
    "ঙ্ক্ষ": "•¶",
    "ঙ্খ": "•L",
    "ঙ্গ": "½",
    "ঙ্ঘ": "•N",

    "চ্ছ্ব": "”Q¡",
    "চ্ছ্ব": "”Q¦",
    "চ্চ": "”P",
    "চ্ছ": "”Q",
    "চ্ঞ": "”T",
    "জ্জ্ব": "¾¡",
    "জ্জ": "¾",
    "জ্ঝ": "À",
    "জ্ঞ": "Á",
    "জ্ব": "R¡",
    "ঞ্চ": "Â",
    "ঞ্ছ": "Ã",
    "ঞ্জ": "Ä",
    "ঞ্ঝ": "Å",
    "ট্ট": "Æ",
    "ট্ব": "U¡",
    "ট্ম": "U¥",
    "ড্ড": "Ç",
    "ণ্ট": "È",
    "ণ্ঠ": "É",
    "ন্স": "Ý",
    "ণ্ড": "Ê",
    "ন্তু": "š‘",
    "ণ্ব": "Y\\^",
    "ত্ত্ব": "Ë¡",
    "ত্ত": "Ë",
    "ত্থ": "Ì",
    "ত্ম": "Z¥",
    "ন্ত্ব": "š—¡",
    "ত্ব": "Z¡",
    "ত্র": "Î",
    "থ্ব": "_¡",
    "ন্দ্ব": "›Ø",
    "দ্গ": "˜M",
    "দ্ঘ": "˜N",
    "দ্দ": "Ï",
    "দ্ধ": "×",
    "দ্ব": "˜¡",
    "দ্ব": "Ø",
    "দ্ভ": "™¢",
    "দ্ম": "Ù",
    "দ্রু": "`ª“",
    "ধ্ব": "aŸ",
    "ধ্ম": "a¥",
    "ন্ট": "›U",
    "ন্ঠ": "Ú",
    "ন্ড": "Û",
    "ন্ত": "šÍ",
    "ন্ত্র": "š¿",
    "ন্থ": "š’",
    "ন্দ": "›`",
    "ন্ধ": "Ü",
    "ন্ন": "bœ",
    "ন্ব": "š\\^",
    "ন্ম": "b¥",
    "প্ট": "Þ",
    "প্ত": "ß",
    "প্ন": "cœ",
    "প্প": "à",
    "প্ল": "cø",
    "প্ল": "c­",
    "প্স": "á",
    "ফ্ল": "d¬",
    "ব্জ": "â",
    "ব্দ": "ã",
    "ব্ধ": "ä",
    "ব্ব": "eŸ",
    "ব্ল": "e­",
    "ব্ল": "eø",
    "ভ্রু": "å“",
    "ভ্র": "å",
    "ম্ন": "gœ",
    "ম্প": "¤ú",
    "ম্ফ": "ç",
    "ম্ব": "¤\\^",
    "ম্ভ": "¤¢",
    "ম্ভ্র": "¤£",
    "ম্ম": "¤§",
    "ম্ল": "¤­",
    "ম্ল": "¤ø",
    "রু": "i“",
    "রু": "iæ",
    "রূ": "iƒ",
    "ল্ক": "é",
    "ল্গ": "ê",
    "ল্ট": "ë",
    "ল্ড": "ì",
    "ল্প": "í",
    "ল্ফ": "î",
    "ল্ব": "j¦",
    "ল্ম": "j¥",
    "ল্ল": "j­",
    "ল্ল": "jø",
    "শু": "ï",
    "শ্চ": "ð",
    "শ্ন": "kœ",
    "শ্ব": "k¦",
    "শ্ম": "k¥",
    "শ্ল": "k­",
    "শ্ল": "kø",
    "ষ্ক": "®‹",
    "ষ্ক্র": "®Œ",
    "ষ্ট": "ó",
    "ষ্ঠ": "ô",
    "ষ্ণ": "ò",
    "ষ্প": "®ú",
    "ষ্ফ": "õ",
    "ষ্ম": "®§",
    "স্ক": "¯‹",
    "স্ট": "÷",
    "স্খ": "ö",
    "স্ত": "¯—",
    "স্ত": "¯Í",
    "স্তু": "¯‘",
    "স্ত্র": "¯¿",
    "স্থ": "¯’",
    "স্ন": "mœ",
    "স্প": "¯ú",
    "স্ফ": "ù",
    "স্ব": "¯\\^",
    "স্ম": "¯§",
    "স্ল": "¯­",
    "স্ল": "¯ø",
    "হু": "û",
    "হ্ণ": "nè",
    "হ্ব": "nŸ",
    "হ্ন": "ý",
    "হ্ম": "þ",
    "হ্ল": "n¬",
    "হৃ": "ü",
    "র্": "©",
    "আ": "Av",
    "অ": "A",
    "ই": "B",
    "ঈ": "C",
    "উ": "D",
    "ঊ": "E",
    "ঋ": "F",
    "এ": "G",
    "ঐ": "H",
    "ও": "I",
    "ঔ": "J",
    "ক": "K",
    "খ": "L",
    "গ": "M",
    "ঘ": "N",
    "ঙ": "O",
    "চ": "P",
    "ছ": "Q",
    "জ": "R",
    "ঝ": "S",
    "ঞ": "T",
    "ট": "U",
    "ঠ": "V",
    "ড": "W",
    "ঢ": "X",
    "ণ": "Y",
    "ত": "Z",
    "থ": "_",
    "দ": "`",
    "ধ": "a",
    "ন": "b",
    "প": "c",
    "ফ": "d",
    "ব": "e",
    "ভ": "f",
    "ম": "g",
    "য": "h",
    "র": "i",
    "ল": "j",
    "শ": "k",
    "ষ": "l",
    "স": "m",
    "হ": "n",
    "ড়": "o",
    "ঢ়": "p",
    "য়": "q",
    "ৎ": "r",
    "০": "0",
    "১": "1",
    "২": "2",
    "৩": "3",
    "৪": "4",
    "৫": "5",
    "৬": "6",
    "৭": "7",
    "৮": "8",
    "৯": "9",
    "া": "v",
    "ি": "w",
    "ী": "x",
    "ু": "æ",
    "ূ": "~",
    "ূ": "‚",
    "ৃ": "„",
    "ে": "†",
    "ৈ": "\\ˆ",
    "ৗ": "Š",
    "-": "Ð",
    "‘": "Ô",
    "’": "Õ",
    "।": "\\|",
    "॥": "\\\\",
    "“": "Ò",
    "”": "Ó",
    "ং": "s",
    "ঃ": "t",
    "ঁ": "u",
    "্র": "ª",
    "্র": "Ö",
    "্র": "«",
    "্য": "¨",
    "্": "\\&",
    "ৃ": "…"
};
// end bijoy_string_conversion_map
/******************************************/



/******************************************
	Rearranges the folas, kars in a 
	unicode string already mapped 
	from ASCII.

	\param str The unicode string

	Coded by : S M Mahbub Murshed
	Date: September 05, 2006
******************************************/

function ReArrangeBijoyClassicConvertedText(str, skipRef) {
    for (var i = 0; i < str.length; i++) {

        // 1. 'Vowel + HALANT + Consonant' -> 'HALANT + Consonant + Vowel'
        if (i > 0 && str.charAt(i) == '\u09CD' &&
            (common.IsBanglaKar(str.charAt(i - 1)) || common.IsBanglaNukta(str.charAt(i - 1))) &&
            i < str.length - 1) {

            var temp = str.substring(0, i - 1);
            temp += str.charAt(i);
            temp += str.charAt(i + 1);
            temp += str.charAt(i - 1);
            temp += str.substring(i + 2);
            str = temp;
        }

        // 2. 'RA + HALANT + Vowel' -> 'Vowel + RA + HALANT'
        if (i > 0 && i < str.length - 1 &&
            str.charAt(i) == '\u09CD' &&
            str.charAt(i - 1) == '\u09B0' &&
            str.charAt(i - 2) != '\u09CD' &&
            common.IsBanglaKar(str.charAt(i + 1))) {

            var temp = str.substring(0, i - 1);
            temp += str.charAt(i + 1);
            temp += str.charAt(i - 1);
            temp += str.charAt(i);
            temp += str.substring(i + 2);
            str = temp;
        }

        // 3. Change Refs (if skipRef is false)
        if (!skipRef) {
            if (i < str.length - 1 && str.charAt(i) == 'র' &&
                common.IsBanglaHalant(str.charAt(i + 1)) &&
                !common.IsBanglaHalant(str.charAt(i - 1))) {

                var j = 1;
                while (true) {
                    if (i - j < 0) break;
                    if (common.IsBanglaBanjonborno(str.charAt(i - j)) &&
                        common.IsBanglaHalant(str.charAt(i - j - 1)))
                        j += 2;
                    else if (j == 1 && common.IsBanglaKar(str.charAt(i - j)))
                        j++;
                    else break;
                }
                var temp = str.substring(0, i - j);
                temp += str.charAt(i);
                temp += str.charAt(i + 1);
                temp += str.substring(i - j, i);
                temp += str.substring(i + 2);
                str = temp;
                i += 1;
                continue;
            }
        }

        // 4. Change pre-kar only if followed by a consonant cluster
        if (i < str.length - 1 &&
            common.IsBanglaPreKar(str.charAt(i)) &&
            !common.IsSpace(str.charAt(i + 1))) {

            // Check if the next consonant has a halant (cluster)
            if (common.IsBanglaBanjonborno(str.charAt(i + 1)) &&
                common.IsBanglaHalant(str.charAt(i + 2))) {

                var temp = str.substring(0, i);
                var j = 1;

                while (common.IsBanglaBanjonborno(str.charAt(i + j))) {
                    if (common.IsBanglaHalant(str.charAt(i + j + 1))) j += 2;
                    else break;
                }

                temp += str.substring(i + 1, i + j + 1);

                var l = 0;
                if (str.charAt(i) == 'ে' && str.charAt(i + j + 1) == 'া') { temp += "ো"; l = 1; }
                else if (str.charAt(i) == 'ে' && str.charAt(i + j + 1) == "ৗ") { temp += "ৌ"; l = 1; }
                else temp += str.charAt(i);

                temp += str.substring(i + j + l + 1);
                str = temp;
                i += j;
            }
            // else: skip pre-kar movement for simple consonant + vowel
        }

        // 5. Nukta after kars
        if (i < str.length - 1 && str.charAt(i) == 'ঁ' &&
            common.IsBanglaPostKar(str.charAt(i + 1))) {

            var temp = str.substring(0, i);
            temp += str.charAt(i + 1);
            temp += str.charAt(i);
            temp += str.substring(i + 2);
            str = temp;
        }
    }

    return str;
}


function ConvertToBijoy(ConvertFrom, line)
{
	var conversion_map = bijoy_string_conversion_map;
	
	for (var ascii in conversion_map)
	{
		var myRegExp = new RegExp(ascii, "g");
		line = line.replace(myRegExp, conversion_map[ascii]);
	}

	var myRegExp = new RegExp("অা", "g");
	line = line.replace(myRegExp, "আ");

	if(ConvertFrom=="bangsee" || ConvertFrom=="bornosoft" ) {
		myRegExp = new RegExp("্্", "g");
		line = line.replace(myRegExp, "্");
	}

	line=ReArrangeBijoyClassicConvertedText(line, true);

	return line;
};
