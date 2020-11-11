# Open Source eDiscovery (Python)
## Create Productions 
Process PDFs, Emails & Attachments (.msg, .eml), Word Docs (.doc, .docx), Excel (.xls,.xlsx), Images (.jpeg, .jpg, .png, .gif) , Video (.mov, .mp4), Text Files (.txt), HTML (.html), Powerpoint (.pptx, .ppt).
Outputs Images (.pdf with Bates Stamp and Image), Natives (.xlsx,.xls, .gif), and Text files (.txt file of extracted text).
Customize folder, batesstamps, and file PREFIXs
Reads collected files from 'input' folder, and Outputs Index Files (.opt and .dat), Images (.jpg with Bates Stamp and Image), Natives (.xlsx,.xls, .gif), and Text files (.txt file of extracted text) in eDiscovery directory stucture (ex.Open-Source_eDiscovery\JSCO_PROD001\VOL0001\NATIVES\JSCO_00000001.xls).

### OUTPUT EXAMPLE
    Example: Create Production from 3 files (.docx, .txt, and .xls) in directory input\Test_Collection
    >Open-Source-eDiscovery
        >JSCO_PROD001
            >VOL00001
                >IMAGES
                    >JSCO00000001.jpg
                    >JSCO00000002.jpg
                >NATIVES
                    >JSCO00000002.xls
                >TEXT
                    >JSCO00000001.txt
                    >JSCO00000002.txt

### How To use - with Python 3.8.2
1. Clone and cd into Open-Source-eDiscovery
1. Place your entire production discovery in the 'input folder'. Make sure any compressed files are extracted (.zips are ignored)
2. Open up email_parser.py and fill in 'User Prompts' with appropriate values.
3. You must also save individual .eml as a pdf by opening and clicking 'Save As' or 'Print' using 'Print as PDF' and move the copies to temp/EMAILS with same name as .elm.
4. pip install -r requirements.txt
5. Run 'python email_parser'


## Search Productions
Search eDicovery Production Index Files, Native Files, and TXT Files. 
Keyword Searching (Email Addresses (Email Fields: From, To, Bcc, Cc), Extracted Text)
Date Range Searching ('Sent On', and 'File Last Modified').

### OUTPUT EXAMPLE
    Example: Search JSCO_PROD0002 for keywords "Lorem" and "lorem.Ipsum@gmail.com" after 11/10/2020
    >Open-Source-eDiscovery
        >SearchProduction
            >RESULTS_002_LOREM_dykwb
                >JSCO00000001.txt
                >JSCO00000001.jpg
                >JSCO00000002.txt
                >JSCO00000002.jpg

### How To use - with Python 3.8.2
1. Clone and cd into Open-Source-eDiscovery\SearchProduction
2. Open parse_dat.py and update 'User Prompts'
3. Run 'python parse_dat'

## Version Changelog

== v1.0.0 ==
Create Production
Search Production

## License

 == MIT ==
 Support Open Source Software
