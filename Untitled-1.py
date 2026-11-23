
import json, os, random, re, secrets, sqlite3, time
from functools import wraps
from urllib.parse import quote
from flask import Flask, Response, g, redirect, render_template_string, request, session
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "isko-badal-do-123")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "acc_data.db")
EXAM_SECONDS = 300
SHOW_OTP_ON_SCREEN = True  # asli SMS jodne ke baad False karo
SEED_VERSION = 2  # jab bhi neeche ke notes/sawal badlo, ye number badha do (3, 4...) to database update ho jayega

# ------------------------------------------------------------------ COURSES (notes English + Hindi)
SEED = {
"basic": ("Basic Computer",
 ["About the course: Basic Computer is a beginner course of about 2 months. You learn what a computer is, how to use Windows, MS Word, Excel, PowerPoint, the internet and email. No earlier knowledge is needed.",
  "A computer takes input, processes it and gives output. Input devices: keyboard, mouse, scanner. Output devices: monitor, printer.",
  "CPU (Central Processing Unit) is the brain of the computer. RAM is temporary memory; data in RAM is lost when the computer is switched off. The hard disk stores data permanently. When you switch on, the BIOS checks the hardware and starts loading the operating system.",
  "Windows is an operating system. It gives the desktop, icons and taskbar. Deleted files first go to the Recycle Bin. Linux is another operating system; an OS is called system software.",
  "Shortcuts: Ctrl+C copy, Ctrl+V paste, Ctrl+X cut, Ctrl+S save, Ctrl+Z undo, Ctrl+A select all, Ctrl+P print.",
  "Memory units: 1 KB = 1024 bytes, 1 MB = 1024 KB, 1 GB = 1024 MB. Computers use the binary system, which has only two digits: 0 and 1.",
  "The internet connects computers worldwide. A browser such as Chrome opens websites. A web address is called a URL (Uniform Resource Locator). Email is used to send letters and files online."],
 ["कोर्स के बारे में: बेसिक कंप्यूटर शुरुआत करने वालों के लिए लगभग 2 महीने का कोर्स है। इसमें कंप्यूटर क्या है, Windows, MS Word, Excel, PowerPoint, इंटरनेट और ईमेल चलाना सीखते हैं। पहले से कुछ आना ज़रूरी नहीं।",
  "कंप्यूटर इनपुट लेता है, उसे प्रोसेस करता है और आउटपुट देता है। इनपुट डिवाइस: कीबोर्ड, माउस, स्कैनर। आउटपुट डिवाइस: मॉनिटर, प्रिंटर।",
  "CPU (सेंट्रल प्रोसेसिंग यूनिट) कंप्यूटर का दिमाग है। RAM अस्थायी मेमोरी है; कंप्यूटर बंद होते ही RAM का डेटा चला जाता है। हार्ड डिस्क में डेटा हमेशा रहता है। कंप्यूटर चालू करते ही BIOS हार्डवेयर जाँचता है और ऑपरेटिंग सिस्टम चालू करवाता है।",
  "Windows एक ऑपरेटिंग सिस्टम है, जो डेस्कटॉप, आइकन और टास्कबार देता है। डिलीट की हुई फ़ाइल पहले Recycle Bin में जाती है। Linux भी एक ऑपरेटिंग सिस्टम है; OS को सिस्टम सॉफ़्टवेयर कहते हैं।",
  "शॉर्टकट: Ctrl+C कॉपी, Ctrl+V पेस्ट, Ctrl+X कट, Ctrl+S सेव, Ctrl+Z अनडू, Ctrl+A सब चुनना, Ctrl+P प्रिंट।",
  "मेमोरी की इकाइयाँ: 1 KB = 1024 बाइट, 1 MB = 1024 KB, 1 GB = 1024 MB। कंप्यूटर बाइनरी सिस्टम इस्तेमाल करता है, जिसमें सिर्फ़ दो अंक होते हैं: 0 और 1।",
  "इंटरनेट पूरी दुनिया के कंप्यूटरों को जोड़ता है। Chrome जैसा ब्राउज़र वेबसाइट खोलता है। वेबसाइट के पते को URL (यूनिफ़ॉर्म रिसोर्स लोकेटर) कहते हैं। ईमेल से ऑनलाइन चिट्ठी और फ़ाइल भेजी जाती है।"]),
"excel": ("MS Excel",
 ["About the course: MS Excel is a course of about 2 months. You learn sheets, formulas, formatting and charts. It is very useful for office work, shops and accounts.",
  "Excel files use the .xlsx extension. A workbook has many sheets. A sheet has rows (numbers) and columns (letters); a cell address looks like A1. A sheet has up to 1,048,576 rows. Shift+F11 inserts a new sheet.",
  "Every formula starts with =. =SUM adds, =AVERAGE gives the average (the average of 10, 20, 30 is 20), =COUNT counts cells with numbers, =MAX gives the biggest number and =MIN the smallest.",
  "The fill handle (small square at a cell corner) copies formulas and series downwards. An absolute reference like $A$1 does not change when copied.",
  "Formatting: Ctrl+B bold, Ctrl+I italic, Ctrl+U underline; borders, colours and number formats such as currency or date.",
  "Condition formulas: =IF(A1>50,\"Pass\",\"Fail\") gives Pass if the condition is true, otherwise Fail. =COUNTIF(A1:A5,\">10\") counts the cells bigger than 10.",
  "Charts and printing: select the data and use Insert to make a chart. Ctrl+P prints and Ctrl+S saves. Always check Print Preview before printing."],
 ["कोर्स के बारे में: MS Excel लगभग 2 महीने का कोर्स है। इसमें शीट, फ़ॉर्मूला, फ़ॉर्मेटिंग और चार्ट सीखते हैं। ऑफिस, दुकान और हिसाब-किताब के काम में बहुत काम आता है।",
  "Excel फ़ाइल का एक्सटेंशन .xlsx होता है। एक वर्कबुक में कई शीट होती हैं। शीट में rows (नंबर) और columns (अक्षर) होते हैं; सेल का पता A1 जैसा होता है। एक शीट में 1,048,576 rows तक होती हैं। Shift+F11 से नई शीट जुड़ती है।",
  "हर फ़ॉर्मूला = से शुरू होता है। =SUM जोड़ता है, =AVERAGE औसत देता है (10, 20, 30 का औसत 20), =COUNT नंबर वाले सेल गिनता है, =MAX सबसे बड़ा और =MIN सबसे छोटा नंबर देता है।",
  "Fill handle (सेल के कोने का छोटा चौकोर) फ़ॉर्मूला और सीरीज़ नीचे तक कॉपी करता है। Absolute reference $A$1 कॉपी करने पर नहीं बदलता।",
  "फ़ॉर्मेटिंग: Ctrl+B बोल्ड, Ctrl+I इटैलिक, Ctrl+U अंडरलाइन; बॉर्डर, रंग और करेंसी या तारीख जैसे नंबर फ़ॉर्मेट।",
  "शर्त वाले फ़ॉर्मूले: =IF(A1>50,\"Pass\",\"Fail\") शर्त सही हो तो Pass, गलत हो तो Fail देता है। =COUNTIF(A1:A5,\">10\") उन सेल को गिनता है जो 10 से बड़े हैं।",
  "चार्ट और प्रिंट: डेटा चुनकर Insert से चार्ट बनाओ। Ctrl+P प्रिंट और Ctrl+S सेव करता है। प्रिंट से पहले Print Preview ज़रूर देखो।"]),
"adv_excel": ("Advanced Excel",
 ["About the course: Advanced Excel is for people who already know basic Excel (about 2 to 3 months). You learn VLOOKUP, IF with AND/OR, Pivot Tables, Data Validation, charts and dashboards. Very useful for office, accounts and data jobs.",
  "VLOOKUP(lookup value, table, column number, FALSE) looks for the value in the first column of a table and returns data from another column. FALSE means exact match and TRUE means approximate match. Newer Excel also has XLOOKUP.",
  "=IF(A1>=40,\"Pass\",\"Fail\"). AND needs all conditions to be true; OR needs at least one to be true. SUMIF adds the cells that match a condition and COUNTIF counts them.",
  "Make a Pivot Table from Insert > PivotTable. Drag fields to Rows, Columns, Values and Filters. It summarises big data in seconds. Refresh it when the data changes.",
  "Data Validation (drop-down list) limits what can be typed in a cell. Conditional Formatting colours cells by rules. Remove Duplicates deletes repeated rows. Sort and Filter arrange data.",
  "LEFT, RIGHT and MID take a part of the text. & or CONCAT joins text. TODAY() gives today's date. LEN counts the characters.",
  "A column chart compares items, a line chart shows change over time and a pie chart shows the share of a whole. A Slicer filters a Pivot Table with one click."],
 ["कोर्स के बारे में: Advanced Excel उनके लिए है जो बेसिक Excel जानते हैं (लगभग 2-3 महीने)। इसमें VLOOKUP, IF के साथ AND/OR, Pivot Table, Data Validation, चार्ट और डैशबोर्ड सीखते हैं। ऑफिस, अकाउंट और डेटा के काम में बहुत उपयोगी।",
  "VLOOKUP(खोजने की वैल्यू, टेबल, कॉलम नंबर, FALSE) टेबल के पहले कॉलम में वैल्यू ढूँढकर दूसरे कॉलम का डेटा देता है। FALSE का मतलब सटीक मैच, TRUE का मतलब अनुमानित मैच। नए Excel में XLOOKUP भी है।",
  "=IF(A1>=40,\"Pass\",\"Fail\")। AND में सभी शर्तें सही होनी चाहिए; OR में कम से कम एक शर्त सही हो। SUMIF शर्त वाले सेल जोड़ता है और COUNTIF उन्हें गिनता है।",
  "Insert > PivotTable से बनता है। फ़ील्ड को Rows, Columns, Values और Filters में ड्रैग करते हैं। बड़े डेटा का सार सेकंडों में बन जाता है। डेटा बदलने पर Refresh करना पड़ता है।",
  "Data Validation (ड्रॉप-डाउन लिस्ट) सेल में लिखने की सीमा तय करता है। Conditional Formatting नियम के अनुसार सेल को रंग देता है। Remove Duplicates दोहराई गई rows हटाता है। Sort और Filter डेटा सजाते हैं।",
  "LEFT, RIGHT, MID टेक्स्ट का हिस्सा निकालते हैं। & या CONCAT टेक्स्ट जोड़ते हैं। TODAY() आज की तारीख देता है। LEN अक्षर गिनता है।",
  "कॉलम चार्ट चीज़ों की तुलना करता है, लाइन चार्ट समय के साथ बदलाव दिखाता है, पाई चार्ट पूरे में से हिस्सा दिखाता है। Slicer एक क्लिक में Pivot Table को फ़िल्टर करता है।"]),
"graphic": ("Graphic Design",
 ["About the course: Graphic Design is a course of about 3 months. You learn to make logos, posters, social media posts and banners with Photoshop, Canva and CorelDRAW. It helps in freelancing, print shops and jobs.",
  "Graphic design combines images, colour and text to give a message. Photoshop is for photo editing, CorelDRAW and Illustrator are for vector work and logos, and Canva is an easy online design tool.",
  "Raster images (JPG, PNG) are made of pixels and become blurry when enlarged. Vector images (SVG, AI) stay sharp at any size. PNG supports a transparent background and JPG files are smaller.",
  "RGB (red, green, blue) is for screens. CMYK (cyan, magenta, yellow, black) is for printing; the K stands for Key, which means black. A colour code looks like #FF0000.",
  "Layers let you edit parts of a design separately. Ctrl+Z undoes. Resolution is measured in DPI/PPI: 300 DPI for print and 72 PPI for screens. Save as PSD to keep the layers.",
  "Fonts are of two kinds: Serif (small strokes at the ends of letters) and Sans-serif (no such strokes). Good design uses contrast, alignment, spacing and readable fonts.",
  "Common sizes: Instagram post 1080 x 1080 px, YouTube thumbnail 1280 x 720 px, A4 paper 210 x 297 mm."],
 ["कोर्स के बारे में: ग्राफ़िक डिज़ाइन लगभग 3 महीने का कोर्स है। इसमें Photoshop, Canva और CorelDRAW से लोगो, पोस्टर, सोशल मीडिया पोस्ट और बैनर बनाना सीखते हैं। फ्रीलांसिंग, प्रिंटिंग शॉप और नौकरी में काम आता है।",
  "ग्राफ़िक डिज़ाइन में चित्र, रंग और टेक्स्ट मिलाकर संदेश दिया जाता है। Photoshop फ़ोटो एडिटिंग के लिए, CorelDRAW और Illustrator वेक्टर व लोगो के लिए, और Canva आसान ऑनलाइन डिज़ाइन के लिए है।",
  "Raster इमेज (JPG, PNG) पिक्सल से बनती हैं और बड़ा करने पर धुंधली हो जाती हैं। Vector इमेज (SVG, AI) किसी भी साइज़ में साफ़ रहती हैं। PNG में पारदर्शी बैकग्राउंड मिलता है और JPG का साइज़ छोटा होता है।",
  "RGB (लाल, हरा, नीला) स्क्रीन के लिए होता है। CMYK (सायन, मैजेंटा, येलो, ब्लैक) प्रिंटिंग के लिए होता है; इसमें K का मतलब Key यानी ब्लैक है। रंग का कोड #FF0000 जैसा होता है।",
  "Layers से डिज़ाइन के हिस्से अलग-अलग एडिट होते हैं। Ctrl+Z अनडू करता है। रेज़ोल्यूशन DPI/PPI में नापा जाता है: प्रिंट के लिए 300 DPI और स्क्रीन के लिए 72 PPI। Layers के साथ सेव करने के लिए PSD फ़ॉर्मेट चुनो।",
  "फ़ॉन्ट दो तरह के होते हैं: Serif (अक्षर के किनारे पर छोटी लाइन) और Sans-serif (बिना लाइन)। अच्छे डिज़ाइन में कंट्रास्ट, अलाइनमेंट, स्पेस और पढ़ने लायक फ़ॉन्ट होते हैं।",
  "आम साइज़: Instagram पोस्ट 1080 x 1080 px, YouTube थंबनेल 1280 x 720 px, A4 पेपर 210 x 297 mm।"]),
"digital": ("Digital Marketing",
 ["About the course: Digital Marketing is a course of about 3 months. You learn SEO, social media, Google Ads, email/WhatsApp marketing and analytics. Useful for shop owners, freelancers and job seekers.",
  "Digital marketing means promoting a business online. Its ways are SEO, social media, email, Google Ads and content.",
  "SEO helps a website rank higher in Google without paying (organic). A keyword is a word people search in Google. On-page SEO means title, keywords and content; off-page SEO means getting backlinks from other websites.",
  "Google Ads is paid promotion (SEM). In PPC (Pay Per Click) you pay only when someone clicks your ad. An impression means how many times the ad was shown. CTR = clicks / impressions x 100.",
  "Post regularly on Instagram, Facebook and YouTube and make reels. Engagement means likes, comments and shares. Hashtags help more people find you.",
  "Use an email list and newsletter to stay in touch with customers. WhatsApp Business lets you add a catalogue. Always take people's permission before sending and never spam.",
  "Google Analytics shows visitors and bounce rate; bounce rate is the visitors who leave after seeing only one page. A conversion happens when a visitor does what you want, such as buying or calling."],
 ["कोर्स के बारे में: डिजिटल मार्केटिंग लगभग 3 महीने का कोर्स है। इसमें SEO, सोशल मीडिया, Google Ads, ईमेल/WhatsApp मार्केटिंग और एनालिटिक्स सीखते हैं। दुकानदार, फ्रीलांसर और नौकरी चाहने वालों के काम आता है।",
  "डिजिटल मार्केटिंग का मतलब है ऑनलाइन बिज़नेस का प्रचार। इसके तरीके: SEO, सोशल मीडिया, ईमेल, Google Ads और कंटेंट।",
  "SEO से वेबसाइट बिना पैसे दिए (organic) Google में ऊपर आती है। Keyword वह शब्द है जो लोग Google में खोजते हैं। On-page SEO में टाइटल, keyword और कंटेंट आते हैं; Off-page SEO में दूसरी वेबसाइट से backlink मिलते हैं।",
  "Google Ads पेड प्रचार (SEM) है। PPC यानी Pay Per Click में पैसे तभी लगते हैं जब कोई विज्ञापन पर क्लिक करे। Impression यानी विज्ञापन कितनी बार दिखा। CTR = क्लिक / इम्प्रेशन x 100।",
  "Instagram, Facebook और YouTube पर नियमित पोस्ट और रील डालो। Engagement में लाइक, कमेंट और शेयर आते हैं। हैशटैग से ज़्यादा लोगों तक पहुँचते हैं।",
  "ईमेल लिस्ट और न्यूज़लेटर से ग्राहकों से जुड़े रहो। WhatsApp Business में कैटलॉग लगा सकते हो। हमेशा लोगों की इजाज़त लेकर भेजो और स्पैम मत करो।",
  "Google Analytics विज़िटर और bounce rate दिखाता है; bounce rate वे विज़िटर हैं जो एक ही पेज देखकर चले जाते हैं। Conversion तब होता है जब विज़िटर वह काम करे जो आप चाहते हो, जैसे खरीदना या कॉल करना।"]),
}

# ------------------------------------------------------------------ TEST KE SAWAL
# Question: (level e/m/h, sawal, [4 options], sahi_option_0_se, topic, explanation)
SEED_Q = {
"basic": [
 ("e","CPU ka full form kya hai?",["Central Processing Unit","Central Program Unit","Computer Personal Unit","Control Processing Unit"],0,"Hardware","CPU = Central Processing Unit, computer ka dimaag."),
 ("e","Keyboard kis tarah ka device hai?",["Output device","Input device","Memory","Software"],1,"Hardware","Keyboard se data daalte hain, isliye input device."),
 ("e","Monitor kis tarah ka device hai?",["Input device","Storage device","Output device","Software"],2,"Hardware","Monitor result dikhata hai, isliye output device."),
 ("e","Copy ka shortcut kya hai?",["Ctrl+V","Ctrl+C","Ctrl+S","Ctrl+P"],1,"Shortcuts","Ctrl+C copy karta hai."),
 ("e","Paste ka shortcut kya hai?",["Ctrl+V","Ctrl+X","Ctrl+Z","Ctrl+A"],0,"Shortcuts","Ctrl+V paste karta hai."),
 ("m","1 KB kitne byte ke barabar hota hai?",["100","512","1024","2048"],2,"Memory units","1 KB = 1024 byte."),
 ("m","RAM kis tarah ki memory hai?",["Permanent","Temporary","Sirf padhne wali","Cloud"],1,"Memory","RAM temporary hai, band karte hi data chala jata hai."),
 ("m","Windows kya hai?",["Operating system","Browser","Hardware","Printer"],0,"Software","Windows ek operating system hai."),
 ("m","File save karne ka shortcut kya hai?",["Ctrl+P","Ctrl+F","Ctrl+S","Ctrl+N"],2,"Shortcuts","Ctrl+S file save karta hai."),
 ("m","Windows me delete ki hui file pehle kahan jaati hai?",["My Computer","Recycle Bin","Desktop","Downloads"],1,"Windows","Delete ki file pehle Recycle Bin me jaati hai."),
 ("h","1 GB kitne MB ke barabar hota hai?",["1000","1024","512","2048"],1,"Memory units","1 GB = 1024 MB."),
 ("h","Inme se kaunsa system software hai?",["MS Word","Photoshop","Linux","Chrome"],2,"Software","Linux ek operating system hai, yaani system software."),
 ("h","Binary system me kaun se digits hote hain?",["0 aur 1","1 se 9","0 se 9","1 aur 2"],0,"Basics","Binary me sirf 0 aur 1 hote hain."),
 ("h","URL ka full form kya hai?",["Uniform Resource Locator","Universal Real Link","Unique Resource Line","User Resource Locator"],0,"Internet","URL = Uniform Resource Locator."),
 ("h","Computer chalu hote hi hardware check karne ka kaam kaun karta hai?",["BIOS","Browser","Antivirus","Printer"],0,"Hardware","BIOS start par hardware check karke OS load karwata hai."),
],
"excel": [
 ("e","Excel file ka extension kya hota hai?",[".docx",".xlsx",".pptx",".txt"],1,"Basics","Excel ka extension .xlsx hota hai."),
 ("e","Excel me cell ka pata kaisa hota hai?",["A1","1A1","A-1-A","AA"],0,"Basics","Column letter + row number milke A1 banta hai."),
 ("e","Jod (sum) nikalne ka function kaunsa hai?",["=ADD","=TOTAL","=SUM","=PLUS"],2,"Formulas","Jod ke liye =SUM use hota hai."),
 ("e","Excel me formula kis chinh se shuru hota hai?",["#","=","@","$"],1,"Formulas","Formula hamesha = se shuru hota hai."),
 ("e","Excel me columns kis se dikhaye jaate hain?",["Numbers","Letters","Colours","Symbols"],1,"Basics","Columns A, B, C letters se dikhte hain."),
 ("m","=AVERAGE(10,20,30) ka answer kya hoga?",["10","20","30","60"],1,"Formulas","(10+20+30)/3 = 20."),
 ("m","Formula neeche copy karne ke liye kis cheez ka use hota hai?",["Fill handle","Ruler","Zoom","Margin"],0,"Tools","Fill handle kheenchne se formula neeche copy hota hai."),
 ("m","Absolute reference kaisa dikhta hai?",["A1","A$","$A$1","#A1"],2,"References","$ lagane se reference lock ho jata hai."),
 ("m","COUNT function kya karta hai?",["Number wale cells ginta hai","Text jodta hai","Cell hata deta hai","Chart banata hai"],0,"Formulas","COUNT sirf number wale cells ginta hai."),
 ("m","Naya worksheet jodne ka shortcut kya hai?",["Shift+F11","Ctrl+W","Alt+F4","Ctrl+Q"],0,"Shortcuts","Shift+F11 naya sheet jodta hai."),
 ("h","=MAX(5,12,9) ka answer kya hoga?",["5","9","12","26"],2,"Formulas","MAX sabse bada number deta hai, yaani 12."),
 ("h","A1=40 ho to =IF(A1>50,\"Pass\",\"Fail\") kya dega?",["Pass","Fail","40","Error"],1,"Logic","40 > 50 galat hai, isliye Fail."),
 ("h","Text ko bold karne ka shortcut kya hai?",["Ctrl+U","Ctrl+I","Ctrl+K","Ctrl+B"],3,"Formatting","Ctrl+B bold karta hai (Ctrl+I italic, Ctrl+U underline)."),
 ("h","=COUNTIF(A1:A5,\">10\") kya karta hai?",["10 se bade cells ginta hai","10 jodta hai","10 se chhote hatata hai","Average nikalta hai"],0,"Logic","COUNTIF shart wale cells ginta hai."),
 ("h","Ek xlsx sheet me maximum kitni rows hoti hain?",["65,536","1,048,576","10,000","1,00,000"],1,"Basics","xlsx me 1,048,576 rows hoti hain."),
],
"adv_excel": [
 ("e","VLOOKUP kya karta hai?",["Sirf cell jodta hai","Table me value dhoondhkar dusre column ka data deta hai","File save karta hai","Photo badalta hai"],1,"VLOOKUP","VLOOKUP pehle column me value dhoondhkar dusre column se data lata hai."),
 ("e","VLOOKUP me exact match ke liye last argument kya likhte hain?",["TRUE","YES","FALSE","NO"],2,"VLOOKUP","FALSE ka matlab exact match."),
 ("e","Pivot Table kis menu se banta hai?",["Insert > PivotTable","Review > Spelling","View > Zoom","Home > Font"],0,"Pivot Table","Pivot Table Insert menu se banta hai."),
 ("e","Drop-down list banane ka tool kaunsa hai?",["Print Preview","Spell Check","Zoom","Data Validation"],3,"Data tools","Data Validation se cell me drop-down list ya limit lagti hai."),
 ("e","Pie chart kis liye use hota hai?",["Time ke saath badlav dikhane","Poore me se hissa dikhane","Text likhne","Formula banane"],1,"Charts","Pie chart poore ka hissa (share) dikhata hai."),
 ("m","A1=40 ho to =IF(A1>=40,\"Pass\",\"Fail\") kya dega?",["Fail","40","Pass","Error"],2,"Logic",">= ka matlab 40 ya usse zyada, isliye Pass."),
 ("m","AND function kab TRUE deta hai?",["Jab koi ek shart sahi ho","Jab saari shartein sahi hon","Hamesha","Kabhi nahi"],1,"Logic","AND me saari shartein sahi honi chahiye."),
 ("m","SUMIF kya karta hai?",["Cells ginta hai","Cells hata deta hai","Shart wale cells ka jod nikalta hai","Chart banata hai"],2,"Logic","SUMIF sirf shart puri karne wale cells jodta hai."),
 ("m","Repeated rows hatane ka tool kaunsa hai?",["Freeze Panes","Remove Duplicates","Zoom","Merge"],1,"Data tools","Remove Duplicates dohrayi hui rows hata deta hai."),
 ("m","Conditional Formatting kya karta hai?",["File delete karta hai","Formula copy karta hai","Sheet chhupata hai","Niyam ke hisaab se cell ko rang deta hai"],3,"Data tools","Conditional Formatting niyam ke hisaab se cell ka rang badalta hai."),
 ("h","VLOOKUP lookup value ko table ke kis column me dhoondhta hai?",["Aakhri column me","Pehle column me","Kisi bhi column me","Sheet ke bahar"],1,"VLOOKUP","VLOOKUP hamesha table ke pehle column me dhoondhta hai."),
 ("h","Pivot Table me fields kahan drag karte hain?",["Sirf Title me","Sirf Footer me","Rows, Columns, Values aur Filters me","Print area me"],2,"Pivot Table","Fields ko Rows, Columns, Values aur Filters me drag karte hain."),
 ("h","Source data badalne par Pivot Table me kya karna padta hai?",["Delete","Save As","Print","Refresh"],3,"Pivot Table","Data badalne par Pivot Table ko Refresh karna padta hai."),
 ("h","=LEFT(\"COMPUTER\",3) ka answer kya hoga?",["PUT","TER","COM","CO"],2,"Text functions","LEFT baayein se 3 akshar deta hai: COM."),
 ("h","Pivot Table ko ek click me filter karne wala tool kaunsa hai?",["Ruler","Slicer","Margin","Macro"],1,"Charts","Slicer Pivot Table ko ek click me filter karta hai."),
],
"graphic": [
 ("e","Photoshop kis kaam ke liye zyada use hota hai?",["Accounts","Photo editing","Typing","Antivirus"],1,"Tools","Photoshop photo editing ka tool hai."),
 ("e","Canva kya hai?",["Operating system","Hardware","Aasan online design tool","Browser"],2,"Tools","Canva ek aasan online design tool hai."),
 ("e","Screen ke liye kaun sa colour mode hota hai?",["CMYK","Pantone","Gray","RGB"],3,"Colour","Screen ke liye RGB (red, green, blue) use hota hai."),
 ("e","Printing ke liye kaun sa colour mode hota hai?",["CMYK","RGB","HEX","Neon"],0,"Colour","Printing ke liye CMYK use hota hai."),
 ("e","Layers ka fayda kya hai?",["Internet chalta hai","Computer tez hota hai","File delete hoti hai","Design ke hisse alag-alag edit hote hain"],3,"Layers","Layers se design ke hisse alag-alag edit hote hain."),
 ("m","Raster image kis se banti hai?",["Lines","Pixels","Sounds","Fonts"],1,"Images","Raster image pixels se banti hai."),
 ("m","Inme se kaunsa vector format hai?",["JPG","PNG","SVG","BMP"],2,"Images","SVG vector format hai, kisi bhi size me saaf rehta hai."),
 ("m","Transparent (khali) background kis format me milta hai?",["JPG","PNG","DOC","MP3"],1,"Images","PNG me transparent background milta hai."),
 ("m","Print ke liye aam taur par kitna DPI chahiye?",["72","300","10","30"],1,"Print","Print ke liye 300 DPI aur screen ke liye 72 PPI hota hai."),
 ("m","Instagram post ka common size kya hai?",["200 x 200 px","3000 x 500 px","1080 x 1080 px","50 x 50 px"],2,"Sizes","Instagram post aam taur par 1080 x 1080 px hota hai."),
 ("h","Vector image ko bada karne par kya hota hai?",["Dhundhli ho jati hai","Saaf rehti hai","Delete ho jati hai","Rang badal jata hai"],1,"Images","Vector image kisi bhi size me saaf rehti hai."),
 ("h","Layers ke saath Photoshop file save karne ka format kaunsa hai?",["JPG","MP4","TXT","PSD"],3,"Layers","PSD format layers ko surakshit rakhta hai."),
 ("h","CMYK me K ka matlab kya hai?",["Kesari","Black (Key)","Blue","Khaki"],1,"Colour","CMYK me K matlab Key, yaani Black."),
 ("h","A4 paper ka size kya hota hai?",["100 x 100 mm","500 x 700 mm","210 x 297 mm","21 x 29 mm"],2,"Sizes","A4 paper 210 x 297 mm ka hota hai."),
 ("h","Sans-serif font ki pehchan kya hai?",["Akshar ke kinaron par chhoti line hoti hai","Akshar ke kinaron par chhoti line nahi hoti","Font sirf Hindi me hota hai","Font sirf bold hota hai"],1,"Typography","Sans-serif me akshar ke kinare par chhoti line nahi hoti."),
],
"digital": [
 ("e","SEO ka full form kya hai?",["Social Email Offer","Search Engine Optimization","Site Edit Option","Search Event Online"],1,"SEO","SEO = Search Engine Optimization."),
 ("e","Google me bina paise diye upar aane ko kya kehte hain?",["Paid ads","Spam","Organic ranking","Hacking"],2,"SEO","Bina paise diye Google me upar aana organic ranking hai."),
 ("e","Inme se kaunsa social media platform hai?",["Instagram","Notepad","Excel","Paint"],0,"Social media","Instagram ek social media platform hai."),
 ("e","Digital marketing ka matlab kya hai?",["Computer repair","Online business ka prachar","Sirf TV advertisement","Typing"],1,"Basics","Digital marketing ka matlab online business ka prachar hai."),
 ("e","Keyword kya hota hai?",["Computer ka button","Password","Printer","Wo shabd jo log Google me search karte hain"],3,"SEO","Keyword wo shabd hai jo log Google me search karte hain."),
 ("m","PPC ka full form kya hai?",["Print Per Copy","Pay Per Click","Page Per Click","Post Per Comment"],1,"Paid ads","PPC = Pay Per Click."),
 ("m","PPC me paise kab lagte hain?",["Ad banate hi","Roz subah","Jab koi ad par click kare","Kabhi nahi"],2,"Paid ads","PPC me paise tabhi lagte hain jab koi ad par click kare."),
 ("m","Backlink kya hota hai?",["Computer ka wire","Dusri website se aapki website ka link","Email ka naam","Photo ka size"],1,"SEO","Backlink dusri website se aapki website ki taraf aane wala link hai."),
 ("m","Impression ka matlab kya hai?",["Ad par kitne click hue","Kitna paisa laga","Ad kitni baar dikhaya gaya","Kitne log kharide"],2,"Paid ads","Impression yaani ad kitni baar dikha."),
 ("m","Engagement me kya aata hai?",["Sirf password","Sirf bijli","Sirf printer","Like, comment aur share"],3,"Social media","Engagement me like, comment aur share aate hain."),
 ("h","CTR ka formula kya hai?",["Impressions x Clicks","Clicks / Impressions x 100","Paisa / Din","Likes - Comments"],1,"Paid ads","CTR = Clicks / Impressions x 100."),
 ("h","Bounce rate kya batata hai?",["Kitne log kharide","Kitna paisa bacha","Jo visitor ek hi page dekh ke chale gaye","Kitne followers badhe"],2,"Analytics","Bounce rate unn visitors ka hai jo ek hi page dekh ke chale gaye."),
 ("h","Conversion ka matlab kya hai?",["Website ka rang badalna","File convert karna","Post delete karna","Visitor ka wo kaam karna jo aap chahte ho (jaise khareedna ya call karna)"],3,"Analytics","Conversion tab hota hai jab visitor khareedna ya call karna jaisa kaam kare."),
 ("h","On-page SEO me kya aata hai?",["Backlink banana","Title, keywords aur content","Paid ad chalana","Email list banana"],1,"SEO","On-page SEO me title, keywords aur content aate hain."),
 ("h","Email ya WhatsApp marketing me sabse zaruri baat kya hai?",["Sabko bina puche bhejna","Roz 100 message bhejna","Logon ki permission lena aur spam na karna","Fake naam use karna"],2,"Email marketing","Hamesha logon ki permission lekar bhejo aur spam mat karo."),
],
}
DEFAULTS = {"fees": "Basic Computer - Rs 1500 (2 mahine)\nMS Excel - Rs 2000 (2 mahine)\nAdvanced Excel - Rs 2500 (2 mahine)\nGraphic Design - Rs 3000 (3 mahine)\nDigital Marketing - Rs 3000 (3 mahine)\nBatch: subah 8-10 baje aur shaam 5-7 baje",
            "whatsapp": "910000000000", "address": "Apna address yahan likho"}

# ------------------------------------------------------------------ DATABASE
def db():
    if "db" not in g:
        g.db = sqlite3.connect(DB); g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def _close(_):
    if "db" in g: g.db.close()

def q(sql, a=(), one=False):
    r = db().execute(sql, a).fetchall()
    return (r[0] if r else None) if one else r

def ex(sql, a=()):
    c = db().execute(sql, a); db().commit(); return c

def init_db():
    c = sqlite3.connect(DB)
    c.executescript("""
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT, phone TEXT UNIQUE, pw TEXT);
    CREATE TABLE IF NOT EXISTS courses(id TEXT PRIMARY KEY, name TEXT, en TEXT, hi TEXT, video TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS questions(id INTEGER PRIMARY KEY, cid TEXT, lv TEXT, q TEXT, o TEXT, ans INTEGER, topic TEXT, exp TEXT);
    CREATE TABLE IF NOT EXISTS attempts(id INTEGER PRIMARY KEY, uid INTEGER, name TEXT, dob TEXT, cid TEXT, course TEXT,
      phone TEXT, score INTEGER, total INTEGER, pct REAL, taken TEXT, detail TEXT);
    CREATE TABLE IF NOT EXISTS progress(uid INTEGER, cid TEXT, idx INTEGER, PRIMARY KEY(uid,cid,idx));
    CREATE TABLE IF NOT EXISTS reviews(id INTEGER PRIMARY KEY, uid INTEGER, name TEXT, txt TEXT, stars INTEGER, ok INTEGER DEFAULT 0);
    CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);""")
    for k, v in DEFAULTS.items():
        c.execute("INSERT OR IGNORE INTO settings VALUES(?,?)", (k, v))
    r = c.execute("SELECT v FROM settings WHERE k='seed_version'").fetchone()
    if not r or r[0] != str(SEED_VERSION):
        for cid, (name, en, hi) in SEED.items():
            c.execute("INSERT OR IGNORE INTO courses(id,name) VALUES(?,?)", (cid, name))
            c.execute("UPDATE courses SET name=?, en=?, hi=? WHERE id=?", (name, "\n".join(en), "\n".join(hi), cid))
            c.execute("DELETE FROM questions WHERE cid=?", (cid,))
        for cid, qs in SEED_Q.items():
            for lv, t, o, a, tp, e in qs:
                c.execute("INSERT INTO questions(cid,lv,q,o,ans,topic,exp) VALUES(?,?,?,?,?,?,?)",
                          (cid, lv, t, json.dumps(o, ensure_ascii=False), a, tp, e))
        c.execute("INSERT OR REPLACE INTO settings VALUES('seed_version',?)", (str(SEED_VERSION),))
    c.commit(); c.close()

def setting(k):
    r = q("SELECT v FROM settings WHERE k=?", (k,), one=True)
    return r["v"] if r else DEFAULTS.get(k, "")

def ready(cid):
    c = {r["lv"]: r["n"] for r in q("SELECT lv,COUNT(*) n FROM questions WHERE cid=? GROUP BY lv", (cid,))}
    return all(c.get(l, 0) >= 5 for l in "emh")

def yt(url):
    m = re.search(r"(?:v=|youtu\.be/|embed/)([\w-]{11})", url or "")
    return m.group(1) if m else None

@app.template_filter("lines")
def lines(s):
    return [x.strip() for x in (s or "").split("\n") if x.strip()]

# ------------------------------------------------------------------ LAYOUT
BASE = """<!doctype html><html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ title }} | All Computer Course.</title><meta name="description" content="{{ desc or 'Hindi aur English me computer course, notes, free test aur certificate.' }}">
<style>
*{box-sizing:border-box}body{margin:0;font:17px/1.6 system-ui,"Noto Sans Devanagari",sans-serif;background:#f4f7fb;color:#12203a}
header{background:#12203a;color:#fff;padding:12px 18px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
header a{color:#ffd166;text-decoration:none;margin-left:14px}header b a{color:#fff;margin:0;font-size:1.15rem}
main{max-width:760px;margin:auto;padding:22px 16px 60px}h1,h2{line-height:1.2}
.card{background:#fff;border:1px solid #d9e1ec;border-radius:12px;padding:16px;margin:12px 0}
.btn,button{display:inline-block;background:#ffb400;color:#2a1d00;border:0;border-radius:10px;padding:10px 18px;font:inherit;font-weight:600;text-decoration:none;cursor:pointer;margin:3px 3px 3px 0}
.btn.alt{background:#fff;border:2px solid #12203a;color:#12203a}.btn.wa{background:#25d366;color:#fff}
input,select,textarea{width:100%;padding:11px;margin:4px 0 12px;border:1px solid #c5cfdd;border-radius:8px;font:inherit}
label{font-weight:600}.err{background:#fde2e2;color:#8a1c1c;padding:10px;border-radius:8px}.ok{background:#dff5e6;color:#17612f;padding:10px;border-radius:8px}
.warn{background:#fff3cd;padding:10px;border-radius:8px}.row{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
table{width:100%;border-collapse:collapse;font-size:.95rem}td,th{padding:8px;border-bottom:1px solid #d9e1ec;text-align:left}
.opt{display:block;padding:8px 10px;border:1px solid #d9e1ec;border-radius:8px;margin:6px 0;font-weight:400}
.good{background:#dff5e6;border-color:#17612f}.bad{background:#fde2e2;border-color:#8a1c1c}
.bar{height:8px;background:#e3e9f2;border-radius:6px;overflow:hidden;width:100%}.bar i{display:block;height:100%;background:#0f766e}
.vid{position:relative;padding-top:56%}.vid iframe{position:absolute;inset:0;width:100%;height:100%;border:0;border-radius:12px}
#timer{position:sticky;top:0;background:#12203a;color:#ffd166;padding:8px;border-radius:0 0 10px 10px;font-weight:700;text-align:center;z-index:5}
.cert{border:8px double #b8860b;background:#fffdf5;padding:30px 18px;text-align:center}
@media print{header,.noprint{display:none}body{background:#fff}.card{border:0}}
</style></head><body>
<header><b><a href="/go">All Computer Course.</a></b><span>
{% if user %}<a href="/home">Home</a><a href="/test">Test</a><a href="/history">History</a><a href="/reviews">⭐ Review</a><a href="/logout">Logout</a>
{% else %}<a href="/free">Free</a><a href="/fees">Fees</a><a href="/toppers">Toppers</a><a href="/login">Login</a>{% endif %}</span></header>
<main>{{ body|safe }}</main></body></html>"""

def page(title, body, desc="", **ctx):
    return render_template_string(BASE, title=title, desc=desc, body=render_template_string(body, **ctx), user=session.get("uid"))

def login_required(f):
    @wraps(f)
    def w(*a, **k):
        return f(*a, **k) if session.get("uid") else redirect("/login")
    return w

def admin_required(f):
    @wraps(f)
    def w(*a, **k):
        return f(*a, **k) if session.get("admin") else redirect("/admin/login")
    return w

def me():
    return q("SELECT * FROM users WHERE id=?", (session["uid"],), one=True)

PHONE = re.compile(r"^[6-9]\d{9}$")

# ------------------------------------------------------------------ SPLASH + AUTH
@app.route("/")
def splash():
    return """<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>All Computer Course</title>
<style>body{margin:0;height:100vh;display:grid;place-items:center;background:#12203a;color:#fff;font:800 clamp(2rem,9vw,4.5rem)/1.1 system-ui;text-align:center}
small{display:block;font:400 1rem system-ui;color:#ffd166;margin-top:14px}</style><div>All Computer Course<small>Loading...</small></div>
<script>setTimeout(function(){location.href='/go'},5000)</script>"""

@app.route("/go")
def go():
    return redirect("/home" if session.get("uid") else "/login")

LOGIN = """<h1>Login</h1>{% if err %}<p class="err">{{ err }}</p>{% endif %}{% if msg %}<p class="ok">{{ msg }}</p>{% endif %}
<form method="post" class="card"><label>Phone number</label><input name="phone" inputmode="numeric" maxlength="10" required>
<label>Password</label><input name="pw" type="password" required><button>Login</button></form>
<p><a href="/register">Naya account banao</a> | <a href="/forgot">Forward / Password bhool gaye?</a></p>
<p><a class="btn alt" href="/free">Bina login free chapter dekho</a></p>"""

REGISTER = """<h1>Naya account</h1>{% if err %}<p class="err">{{ err }}</p>{% endif %}
<form method="post" class="card"><label>Aapka naam</label><input name="name" required>
<label>Phone number</label><input name="phone" inputmode="numeric" maxlength="10" required>
<label>Apna password banao</label><input name="pw" type="password" minlength="4" required><button>Account banao</button></form>
<p><a href="/login">Pehle se account hai? Login</a></p>"""

@app.route("/login", methods=["GET", "POST"])
def login():
    err = None
    if request.method == "POST":
        u = q("SELECT * FROM users WHERE phone=?", (request.form["phone"].strip(),), one=True)
        if u and check_password_hash(u["pw"], request.form["pw"]):
            session.clear(); session["uid"] = u["id"]; return redirect("/home")
        err = "Phone number ya password galat hai."
    return page("Login", LOGIN, err=err, msg=request.args.get("msg"))

@app.route("/register", methods=["GET", "POST"])
def register():
    err = None
    if request.method == "POST":
        f = request.form; phone = f["phone"].strip()
        if not PHONE.match(phone): err = "Sahi 10 digit ka phone number daalo."
        elif q("SELECT 1 FROM users WHERE phone=?", (phone,), one=True): err = "Is number se account pehle se hai. Login karo."
        else:
            ex("INSERT INTO users(name,phone,pw) VALUES(?,?,?)", (f["name"].strip(), phone, generate_password_hash(f["pw"])))
            return redirect("/login?msg=Account ban gaya, ab login karo.")
    return page("Register", REGISTER, err=err)

FORGOT = """<h1>Naya password</h1>{% if err %}<p class="err">{{ err }}</p>{% endif %}
{% if step==1 %}<form method="post" class="card"><label>Apna phone number</label><input name="phone" inputmode="numeric" maxlength="10" required><button>OTP bhejo</button></form>
{% else %}{% if otp %}<p class="ok">Demo OTP: <b>{{ otp }}</b></p>{% endif %}
<form method="post" class="card"><label>OTP</label><input name="otp" inputmode="numeric" required>
<label>Naya password</label><input name="pw" type="password" minlength="4" required><button>Password badlo</button></form>{% endif %}"""

@app.route("/forgot", methods=["GET", "POST"])
def forgot():
    err = otp = None; step = 1
    if request.method == "POST":
        f = request.form
        if "phone" in f:
            phone = f["phone"].strip()
            if q("SELECT 1 FROM users WHERE phone=?", (phone,), one=True):
                code = str(secrets.randbelow(900000) + 100000)
                session["reset"] = {"phone": phone, "otp": code, "exp": time.time() + 600}
                print(f"[OTP] {phone}: {code}")  # yahan SMS API (Fast2SMS / MSG91) jodo
                step = 2; otp = code if SHOW_OTP_ON_SCREEN else None
            else: err = "Is number se koi account nahi mila."
        else:
            r = session.get("reset"); step = 2
            if not r or time.time() > r["exp"]: err = "OTP expire ho gaya. Dobara try karo."; step = 1
            elif f["otp"].strip() != r["otp"]: err = "OTP galat hai."
            else:
                ex("UPDATE users SET pw=? WHERE phone=?", (generate_password_hash(f["pw"]), r["phone"]))
                session.pop("reset"); return redirect("/login?msg=Password badal gaya, ab login karo.")
    return page("Forgot", FORGOT, err=err, otp=otp, step=step)

@app.route("/logout")
def logout():
    session.clear(); return redirect("/login")

# ------------------------------------------------------------------ HOME + COURSE (har course alag)
HOME = """<h1>All Computer Course.</h1><h2>Hello, {{ u['name'] }} 👋</h2><p>Apna course chuno:</p>
<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:12px">
{% for c in courses %}{% set total = c['en']|lines|length %}{% set d = prog.get(c['id'],0) %}
<a href="/course/{{ c['id'] }}" style="text-decoration:none;color:inherit"><div class="card" style="margin:0;height:100%">
<h3 style="margin:0 0 8px">{{ c['name'] }}</h3><div class="bar"><i style="width:{{ (d*100/total)|int if total else 0 }}%"></i></div>
<small>{{ d }}/{{ total }} chapter padhe</small></div></a>{% endfor %}</div>
<p><a class="btn" href="/reviews#form">⭐ Review do</a><a class="btn alt" href="/history">Meri History</a><a class="btn alt" href="/fees">Fees</a></p>"""

COURSE = """<div class="noprint"><a href="/home">&larr; Saare courses</a></div><h1>{{ c['name'] }}</h1>
<p class="noprint">{% if ok %}<a class="btn" href="/test/{{ c['id'] }}">Test do</a><a class="btn alt" href="/practice/{{ c['id'] }}">Practice</a>
{% else %}<span class="warn">Is course ka test jaldi aayega</span>{% endif %}</p>
{% if vid %}<div class="vid noprint"><iframe src="https://youtu.be/tHByaMnTl8g?si=2cp6ntb35qG_hzUc/{{ vid }}" allowfullscreen></iframe></div>{% endif %}
{% set en = c['en']|lines %}{% set hi = c['hi']|lines %}
{% for t in en %}<div class="card"><b>Chapter {{ loop.index }}</b>{% if loop.index0 in done %} ✅{% endif %}
<p>{{ t }}</p><p>{{ hi[loop.index0] if loop.index0 < hi|length else '' }}</p>
{% if loop.index0 not in done %}<form method="post" action="/read/{{ c['id'] }}/{{ loop.index0 }}" class="noprint"><button>Padh liya</button></form>{% endif %}</div>{% endfor %}
<p class="noprint"><button onclick="window.print()">PDF save karo</button> <a class="btn alt" href="/reviews#form">⭐ Review do</a></p>
{% if hist %}<h2 class="noprint">Is course me meri history</h2><div class="card noprint" style="overflow-x:auto"><table>
<tr><th>Date</th><th>Number</th><th>%</th><th></th></tr>{% for r in hist %}<tr><td>{{ r['taken'] }}</td><td>{{ r['score'] }}/{{ r['total'] }}</td>
<td><b>{{ '%.0f'|format(r['pct']) }}%</b></td><td><a href="/result/{{ r['id'] }}">Dekho</a>{% if r['pct']>=60 %} | <a href="/cert/{{ r['id'] }}">Certificate</a>{% endif %}</td></tr>{% endfor %}
</table></div>{% endif %}"""

@app.route("/home")
@login_required
def home():
    prog = {r["cid"]: r["n"] for r in q("SELECT cid,COUNT(*) n FROM progress WHERE uid=? GROUP BY cid", (session["uid"],))}
    return page("Home", HOME, u=me(), courses=q("SELECT * FROM courses"), prog=prog)

@app.route("/course/<cid>")
@login_required
def course(cid):
    c = q("SELECT * FROM courses WHERE id=?", (cid,), one=True)
    if not c: return redirect("/home")
    done = {r["idx"] for r in q("SELECT idx FROM progress WHERE uid=? AND cid=?", (session["uid"], cid))}
    hist = q("SELECT * FROM attempts WHERE uid=? AND cid=? ORDER BY id DESC", (session["uid"], cid))
    return page(c["name"], COURSE, c=c, vid=yt(c["video"]), done=done, hist=hist, ok=ready(cid))

@app.route("/read/<cid>/<int:i>", methods=["POST"])
@login_required
def read(cid, i):
    ex("INSERT OR IGNORE INTO progress VALUES(?,?,?)", (session["uid"], cid, i))
    return redirect(f"/course/{cid}")

# ------------------------------------------------------------------ TEST / PRACTICE
TESTS = """<h1>Test</h1><p>15 sawal (5 easy, 5 minute.</p>
{% for c in courses %}<div class="card row"><b>{{ c['name'] }}</b>{% if ready[c['id']] %}<span>
<a class="btn" href="/test/{{ c['id'] }}">Test</a><a class="btn alt" href="/practice/{{ c['id'] }}">Practice</a></span>{% else %}<span>Jaldi aayega</span>{% endif %}</div>{% endfor %}"""

DETAIL = """<h1>{{ c['name'] }} Test</h1><form method="post" class="card"><label>Naam</label><input name="name" value="{{ u['name'] }}" required>
<label>Date of birth</label><input name="dob" type="date" required><label>Course</label><input value="{{ c['name'] }}" readonly>
<label>Phone number</label><input name="phone" value="{{ u['phone'] }}" required>
<p>Start dabate hi 5 minute ka timer chalu hoga. Time khatam hone par test apne aap band ho jayega.</p><button>Start</button></form>"""

EXAM = """{% if left is not none %}<div id="timer">Time: <span id="t"></span></div>{% endif %}<h2>{{ c['name'] }}{% if left is none %} (Practice){% endif %}</h2>
<form method="post" id="f">{% for x in qs %}<div class="card"><b>{{ loop.index }}. {{ x['q'] }}</b>
{% for o in x['o'] %}<label class="opt"><input type="radio" name="q{{ x['id'] }}" value="{{ loop.index0 }}" style="width:auto;margin:0 8px 0 0">{{ o }}</label>{% endfor %}</div>{% endfor %}
<button>Submit</button></form>
{% if left is not none %}<script>var left={{ left }},el=document.getElementById('t');
function tick(){var m=Math.floor(left/60),s=left%60;el.textContent=m+':'+(s<10?'0':'')+s;if(left<=0){document.getElementById('f').submit();return}left--;setTimeout(tick,1000)}tick();</script>{% endif %}"""

RESULT = """<h1>Result</h1><div class="card"><p><b>{{ a['name'] }}</b> &middot; {{ a['course'] }}{% if not saved %} (Practice){% endif %}</p>
<h2>{{ a['score'] }} / {{ a['total'] }} &nbsp; ({{ '%.0f'|format(a['pct']) }}%)</h2>
<p>{% if a['pct']>=60 %}Shabash! Aap pass ho. 🎉{% else %}Thoda aur practice karo, notes dobara padho.{% endif %}</p>
{% if weak %}<p class="warn"><b>Ye topics dobara padho:</b> {{ weak|join(', ') }}</p>{% endif %}</div>
<p>{% if saved %}<a class="btn wa" href="{{ share }}" target="_blank">WhatsApp par share karo</a>
{% if a['pct']>=60 %}<a class="btn" href="/cert/{{ a['id'] }}">Certificate lo</a>{% endif %}{% endif %}
<a class="btn alt" href="/reviews#form">⭐ Review do</a><a class="btn alt" href="/history">History</a><a class="btn alt" href="/home">Home</a></p>
<h2>Sahi jawab aur explanation</h2>{% for r in rev %}<div class="card"><b>{{ loop.index }}. {{ r['q'] }}</b>
{% for o in r['o'] %}<div class="opt {{ 'good' if loop.index0==r['ans'] else ('bad' if loop.index0==r['pick'] else '') }}">{{ o }}
{% if loop.index0==r['ans'] %}✔{% elif loop.index0==r['pick'] %}✘ (aapka jawab){% endif %}</div>{% endfor %}<small>{{ r['exp'] }}</small></div>{% endfor %}"""

CERT = """<div class="cert"><h3>Certificate of Completion</h3><p>Ye certificate diya jata hai</p><h1>{{ a['name'] }}</h1>
<p>ne <b>{{ a['course'] }}</b> test me <b>{{ '%.0f'|format(a['pct']) }}%</b> number laaye</p>
<p>Date: {{ a['taken'] }} &nbsp;|&nbsp; No: ACC-{{ '%05d'|format(a['id']) }}</p><p><b>All Computer Co.</b></p></div>
<p class="noprint"><button onclick="window.print()">Print / PDF save</button> <a class="btn alt" href="/history">Back</a></p>"""

def pick_questions(cid):
    out = []
    for lv in "emh":
        ids = [r["id"] for r in q("SELECT id FROM questions WHERE cid=? AND lv=?", (cid, lv))]
        out += random.sample(ids, 5)
    return out

def review(detail):
    ids = [d[0] for d in detail]
    rows = {r["id"]: r for r in q(f"SELECT * FROM questions WHERE id IN ({','.join('?' * len(ids))})", ids)} if ids else {}
    return [{"q": rows[i]["q"], "o": json.loads(rows[i]["o"]), "ans": rows[i]["ans"], "pick": p, "exp": rows[i]["exp"],
             "topic": rows[i]["topic"], "ok": p == rows[i]["ans"]} for i, p in detail if i in rows]

def weak_topics(rev):
    return list(dict.fromkeys(r["topic"] for r in rev if not r["ok"] and r["topic"]))

@app.route("/test")
@login_required
def tests():
    cs = q("SELECT * FROM courses")
    return page("Test", TESTS, courses=cs, ready={c["id"]: ready(c["id"]) for c in cs})

@app.route("/test/<cid>", methods=["GET", "POST"])
@login_required
def test_detail(cid):
    c = q("SELECT * FROM courses WHERE id=?", (cid,), one=True)
    if not c or not ready(cid): return redirect("/test")
    if request.method == "POST":
        f = request.form
        session["exam"] = {"cid": cid, "mode": "test", "start": time.time(), "qs": pick_questions(cid),
                           "name": f["name"].strip(), "dob": f["dob"], "phone": f["phone"].strip()}
        return redirect("/exam")
    return page("Test", DETAIL, c=c, u=me())

@app.route("/practice/<cid>")
@login_required
def practice(cid):
    if not q("SELECT 1 FROM courses WHERE id=?", (cid,), one=True) or not ready(cid): return redirect("/test")
    u = me()
    session["exam"] = {"cid": cid, "mode": "practice", "start": time.time(), "qs": pick_questions(cid),
                       "name": u["name"], "dob": "", "phone": u["phone"]}
    return redirect("/exam")

@app.route("/exam", methods=["GET", "POST"])
@login_required
def exam():
    ex_ = session.get("exam")
    if not ex_: return redirect("/test")
    c = q("SELECT * FROM courses WHERE id=?", (ex_["cid"],), one=True)
    elapsed = time.time() - ex_["start"]; practice_ = ex_["mode"] == "practice"
    if request.method == "POST":
        late = (not practice_) and elapsed > EXAM_SECONDS + 15  # thoda network grace
        detail = []
        for i in ex_["qs"]:
            p = request.form.get(f"q{i}", "")
            detail.append([i, -1 if late or not p.isdigit() else int(p)])
        rev = review(detail); score = sum(r["ok"] for r in rev); total = len(detail); pct = score * 100 / total
        session.pop("exam")
        if practice_:
            a = {"name": ex_["name"], "course": c["name"], "score": score, "total": total, "pct": pct}
            return page("Practice Result", RESULT, a=a, rev=rev, weak=weak_topics(rev), saved=False, share="")
        cur = ex(("INSERT INTO attempts(uid,name,dob,cid,course,phone,score,total,pct,taken,detail) VALUES(?,?,?,?,?,?,?,?,?,?,?)"),
                 (session["uid"], ex_["name"], ex_["dob"], ex_["cid"], c["name"], ex_["phone"], score, total, pct,
                  time.strftime("%d-%m-%Y %H:%M"), json.dumps(detail)))
        return redirect(f"/result/{cur.lastrowid}")
    rows = {r["id"]: r for r in q(f"SELECT * FROM questions WHERE id IN ({','.join('?' * len(ex_['qs']))})", ex_["qs"])}
    qs = [{"id": i, "q": rows[i]["q"], "o": json.loads(rows[i]["o"])} for i in ex_["qs"]]
    left = None if practice_ else max(0, int(EXAM_SECONDS - elapsed))
    return page("Exam", EXAM, c=c, qs=qs, left=left)

@app.route("/result/<int:aid>")
@login_required
def result(aid):
    a = q("SELECT * FROM attempts WHERE id=? AND uid=?", (aid, session["uid"]), one=True)
    if not a: return redirect("/history")
    rev = review(json.loads(a["detail"])); a = dict(a)
    msg = f"Maine {a['course']} test me {a['pct']:.0f}% number laye! Aap bhi try karo: {request.host_url}"
    return page("Result", RESULT, a=a, rev=rev, weak=weak_topics(rev), saved=True, share="https://wa.me/?text=" + quote(msg))

@app.route("/cert/<int:aid>")
@login_required
def cert(aid):
    a = q("SELECT * FROM attempts WHERE id=? AND uid=?", (aid, session["uid"]), one=True)
    if not a or a["pct"] < 60: return redirect("/history")
    return page("Certificate", CERT, a=dict(a))

HIST = """<h1>Meri History</h1>{% if not rows %}<p>Abhi tak koi test nahi diya.</p>{% else %}
<div class="card" style="overflow-x:auto"><table><tr><th>Date</th><th>Course</th><th>Number</th><th>%</th><th></th></tr>
{% for r in rows %}<tr><td>{{ r['taken'] }}</td><td>{{ r['course'] }}</td><td>{{ r['score'] }}/{{ r['total'] }}</td><td><b>{{ '%.0f'|format(r['pct']) }}%</b></td>
<td><a href="/result/{{ r['id'] }}">Dekho</a>{% if r['pct']>=60 %} | <a href="/cert/{{ r['id'] }}">Certificate</a>{% endif %}</td></tr>{% endfor %}</table></div>{% endif %}"""

@app.route("/history")
@login_required
def history():
    return page("History", HIST, rows=q("SELECT * FROM attempts WHERE uid=? ORDER BY id DESC", (session["uid"],)))

# ------------------------------------------------------------------ PUBLIC PAGES (free, fees, toppers, reviews, SEO)
FREE = """<h1>Free computer notes</h1><p>Hindi aur English me. Poore notes, test aur certificate ke liye account banao.</p>
{% for c in courses %}<div class="card row"><b>{{ c['name'] }}</b><a class="btn alt" href="/free/{{ c['id'] }}">Free chapter padho</a></div>{% endfor %}
<a class="btn" href="/register">Account banao</a>"""

FREE1 = """<h1>{{ c['name'] }} Hindi me seekho (Free Chapter 1)</h1>{% set en=c['en']|lines %}{% set hi=c['hi']|lines %}
<div class="card"><p>{{ en[0] if en else '' }}</p><p>{{ hi[0] if hi else '' }}</p></div>
<p>Aage ke chapter, video, practice test aur certificate ke liye account banao.</p>
<a class="btn" href="/register">Free account banao</a><a class="btn alt" href="/free">Baaki courses</a>"""

FEES = """<h1>Fees aur Batch</h1><div class="card">{% for l in fees|lines %}<p>{{ l }}</p>{% endfor %}<p><b>Address:</b> {{ addr }}</p></div>
<a class="btn wa" href="https://wa.me/{{ wa }}?text={{ 'Namaste, mujhe admission ki jankari chahiye.'|urlencode }}" target="_blank">WhatsApp par pucho</a>
<a class="btn alt" href="tel:+{{ wa }}">Call karo</a>"""

TOP = """<h1>Toppers 🏆</h1>{% if not rows %}<p>Abhi koi topper nahi. Pehle aap ban sakte ho!</p>{% else %}<div class="card"><table>
<tr><th>#</th><th>Naam</th><th>Course</th><th>%</th></tr>{% for r in rows %}<tr><td>{{ loop.index }}</td><td>{{ r['name'].split()[0] }}</td><td>{{ r['course'] }}</td>
<td><b>{{ '%.0f'|format(r['p']) }}%</b></td></tr>{% endfor %}</table></div>{% endif %}"""

REVIEWS = """<h1>Students ke review</h1>{% if msg %}<p class="ok">{{ msg }}</p>{% endif %}
{% if user %}<form method="post" class="card" id="form"><h3 style="margin-top:0">⭐ Apna review likho</h3>
<label>Aapka review</label><textarea name="txt" rows="3" required maxlength="300" placeholder="Course kaisa laga?"></textarea>
<label>Stars</label><select name="stars"><option value="5">5 - Bahut accha</option><option value="4">4 - Accha</option><option value="3">3 - Theek hai</option><option value="2">2</option><option value="1">1</option></select>
<button>Review bhejo</button><p><small>Admin approve karega tab sabko dikhega.</small></p></form>
{% else %}<p class="card"><a href="/login">Review dene ke liye pehle login karo</a></p>{% endif %}
<h2>Sabke review</h2>{% for r in rows %}<div class="card"><b>{{ r['name'].split()[0] }}</b> {{ '⭐' * r['stars'] }}<p>{{ r['txt'] }}</p></div>{% else %}<p>Abhi koi review nahi.</p>{% endfor %}"""

@app.route("/free")
def free():
    return page("Free Computer Notes Hindi me", FREE, "Basic Computer, Excel, Graphic Design aur Digital Marketing ke free notes Hindi aur English me.", courses=q("SELECT * FROM courses"))

@app.route("/free/<cid>")
def free_one(cid):
    c = q("SELECT * FROM courses WHERE id=?", (cid,), one=True)
    if not c: return redirect("/free")
    return page(f"{c['name']} Hindi me seekho", FREE1, f"{c['name']} ka free pehla chapter Hindi aur English me.", c=c)

@app.route("/fees")
def fees():
    return page("Fees aur Batch", FEES, fees=setting("fees"), addr=setting("address"), wa=setting("whatsapp"))

@app.route("/toppers")
def toppers():
    rows = q("SELECT name, course, MAX(pct) p FROM attempts GROUP BY uid, cid ORDER BY p DESC LIMIT 10")
    return page("Toppers", TOP, rows=rows)

@app.route("/reviews", methods=["GET", "POST"])
def reviews():
    msg = None
    if request.method == "POST" and session.get("uid"):
        ex("INSERT INTO reviews(uid,name,txt,stars) VALUES(?,?,?,?)",
           (session["uid"], me()["name"], request.form["txt"][:300], min(5, max(1, int(request.form["stars"])))))
        msg = "Dhanyavaad! Aapka review admin ke approve karne ke baad dikhega."
    return page("Reviews", REVIEWS, rows=q("SELECT * FROM reviews WHERE ok=1 ORDER BY id DESC"), msg=msg)

@app.route("/robots.txt")
def robots():
    return Response(f"User-agent: *\nAllow: /\nSitemap: {request.host_url}sitemap.xml\n", mimetype="text/plain")

@app.route("/sitemap.xml")
def sitemap():
    urls = ["", "free", "fees", "toppers", "reviews"] + [f"free/{c['id']}" for c in q("SELECT id FROM courses")]
    x = "".join(f"<url><loc>{request.host_url}{u}</loc></url>" for u in urls)
    return Response(f'<?xml version="1.0" encoding="UTF-8"?><urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">{x}</urlset>', mimetype="application/xml")

# ------------------------------------------------------------------ ADMIN PANEL
ALOGIN = """<h1>Admin Login</h1>{% if err %}<p class="err">{{ err }}</p>{% endif %}
<form method="post" class="card"><label>Admin password</label><input name="pw" type="password" required><button>Login</button></form>"""

ADMIN = """<h1>Admin Panel</h1><p>Students: <b>{{ n_users }}</b> | Tests diye gaye: <b>{{ n_att }}</b> | <a href="/admin/logout">Admin logout</a></p>{% if msg %}<p class="ok">{{ msg }}</p>{% endif %}
<h2>Fees, WhatsApp aur Address</h2><form method="post" action="/admin/settings" class="card"><label>Fees / batch (har line alag)</label>
<textarea name="fees" rows="5">{{ s['fees'] }}</textarea><label>WhatsApp number (91 ke saath, bina +)</label><input name="whatsapp" value="{{ s['whatsapp'] }}">
<label>Address</label><input name="address" value="{{ s['address'] }}"><button>Save</button></form>
<h2>Naya course jodo</h2><form method="post" action="/admin/course" class="card"><label>Course ka naam</label><input name="name" required>
<label>English notes (har chapter alag line)</label><textarea name="en" rows="4" required></textarea>
<label>Hindi notes (har chapter alag line, usi order me)</label><textarea name="hi" rows="4" required></textarea>
<label>YouTube video link (optional)</label><input name="video"><button>Course jodo</button></form>
<h2>Courses aur video</h2>{% for c in courses %}<form method="post" action="/admin/video/{{ c['id'] }}" class="card"><b>{{ c['name'] }}</b>
<small>(sawal: {{ counts.get(c['id'], 'e0 m0 h0') }})</small><input name="video" value="{{ c['video'] }}" placeholder="YouTube link"><button>Video save</button></form>{% endfor %}
<h2>Naya sawal jodo</h2><form method="post" action="/admin/question" class="card"><label>Course</label><select name="cid">{% for c in courses %}<option value="{{ c['id'] }}">{{ c['name'] }}</option>{% endfor %}</select>
<label>Level</label><select name="lv"><option value="e">Easy</option><option value="m">Normal</option><option value="h">Hard</option></select>
<label>Sawal</label><input name="q" required><label>Option 1</label><input name="o1" required><label>Option 2</label><input name="o2" required>
<label>Option 3</label><input name="o3" required><label>Option 4</label><input name="o4" required>
<label>Sahi option number (1-4)</label><select name="ans"><option>1</option><option>2</option><option>3</option><option>4</option></select>
<label>Topic</label><input name="topic" required><label>Explanation</label><input name="exp" required><button>Sawal jodo</button></form>
<h2>Review approve karo</h2>{% for r in pending %}<div class="card"><b>{{ r['name'] }}</b> {{ '⭐' * r['stars'] }}<p>{{ r['txt'] }}</p>
<form method="post" action="/admin/review/{{ r['id'] }}/ok" style="display:inline"><button>Approve</button></form>
<form method="post" action="/admin/review/{{ r['id'] }}/del" style="display:inline"><button class="btn alt">Delete</button></form></div>{% else %}<p>Koi pending review nahi.</p>{% endfor %}"""

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    err = None
    if request.method == "POST":
        if secrets.compare_digest(request.form["pw"], ADMIN_PASSWORD):
            session["admin"] = True; return redirect("/admin")
        err = "Password galat hai."
    return page("Admin", ALOGIN, err=err)

@app.route("/admin/logout")
def admin_logout():
    session.pop("admin", None); return redirect("/")

@app.route("/admin")
@admin_required
def admin():
    counts = {}
    for r in q("SELECT cid, lv, COUNT(*) n FROM questions GROUP BY cid, lv ORDER BY lv"):
        counts.setdefault(r["cid"], {})[r["lv"]] = r["n"]
    counts = {k: " ".join(f"{l}{v.get(l, 0)}" for l in "emh") for k, v in counts.items()}
    return page("Admin", ADMIN, n_users=q("SELECT COUNT(*) n FROM users", one=True)["n"], n_att=q("SELECT COUNT(*) n FROM attempts", one=True)["n"],
                s={k: setting(k) for k in DEFAULTS}, courses=q("SELECT * FROM courses"), counts=counts,
                pending=q("SELECT * FROM reviews WHERE ok=0"), msg=request.args.get("msg"))

@app.route("/admin/settings", methods=["POST"])
@admin_required
def admin_settings():
    for k in DEFAULTS: ex("INSERT OR REPLACE INTO settings VALUES(?,?)", (k, request.form[k]))
    return redirect("/admin?msg=Save ho gaya")

@app.route("/admin/course", methods=["POST"])
@admin_required
def admin_course():
    f = request.form; cid = re.sub(r"[^a-z0-9]+", "_", f["name"].lower()).strip("_") or secrets.token_hex(3)
    ex("INSERT OR REPLACE INTO courses(id,name,en,hi,video) VALUES(?,?,?,?,?)", (cid, f["name"].strip(), f["en"], f["hi"], f["video"].strip()))
    return redirect("/admin?msg=Course ban gaya. Ab uske 15 sawal jodo (5 easy, 5 normal, 5 hard).")

@app.route("/admin/video/<cid>", methods=["POST"])
@admin_required
def admin_video(cid):
    ex("UPDATE courses SET video=? WHERE id=?", (request.form["video"].strip(), cid)); return redirect("/admin?msg=Video save ho gaya")

@app.route("/admin/question", methods=["POST"])
@admin_required
def admin_question():
    f = request.form
    ex("INSERT INTO questions(cid,lv,q,o,ans,topic,exp) VALUES(?,?,?,?,?,?,?)",
       (f["cid"], f["lv"], f["q"].strip(), json.dumps([f["o1"], f["o2"], f["o3"], f["o4"]], ensure_ascii=False), int(f["ans"]) - 1, f["topic"].strip(), f["exp"].strip()))
    return redirect("/admin?msg=Sawal jud gaya")

@app.route("/admin/review/<int:rid>/<act>", methods=["POST"])
@admin_required
def admin_review(rid, act):
    ex("UPDATE reviews SET ok=1 WHERE id=?" if act == "ok" else "DELETE FROM reviews WHERE id=?", (rid,)); return redirect("/admin")

init_db()
if __name__ == "__main__":
    app.run(debug=True)