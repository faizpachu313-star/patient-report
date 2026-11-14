"""Add the 'RADIOLOGY DATA - OUT PATIENTS' section (Oct-2026 model) under every month sheet.

Data transcribed from the Spectrum Center referral sheets / radiology screenshots.
Styles are copied from the Oct-2026 sheet, which is the design-system source of truth.

Usage: python scripts/add_radiology_sections.py <master.xlsx>
"""
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

# (dd.mm.yyyy, file no, name, examination, bill AED[, note])
# Names cut off in the source are shortened to first + second name (repo rule).
RADIOLOGY = {
    "Nov-2025": [
        ("24.11.2025", "506431213", "SYED MD AWLAD", "CT", 433),
    ],
    "Dec-2025": [
        ("22.12.2025", "TMC092095", "MALIK JAN DILBAR KHAN", "MRI", 782),
        ("26.12.2025", "TMC125661", "NIJU KALLATT AYYAPPAKUTTY AYYAPPAKUTTY", "MRI", 782),
        ("24.12.2025", "TMC125594", "MUHAMMAD RAMZAN MUHAMMAD YOUNAS", "MRI", 782),
        ("06.12.2025", "TMC124715", "SHATRUDHAN SINGH PRABHANSH SINGH", "MRI", 854),
        ("07.12.2025", "TMC124762", "RINA NURHAYATI", "MRI", 835),
        ("06.12.2025", "TMC124708", "RAFID KHAN SAIF UR REHMAN", "MRI", 782),
        ("07.12.2025", "TMC124773", "FUAD ALI PULLOT SHOUKATH ALI PULLOT", "MRI", 782),
        ("10.12.2025", "TMC124961", "AKRAM NAGI NAKHLA IBRAHIM", "MRI", 854),
        ("13.12.2025", "TMC125068", "SHAH MUHAMMAD KHAN MUHAMMAD", "MRI", 854),
        ("13.12.2025", "TMC125104", "MUHAMMAD QAMAR MEHMOOD KHAN MUHAMMAD", "MRI", 782),
        ("17.12.2025", "TMC125293", "SHAHBAZ ALI ASGHAR ALI", "MRI", 795),
        ("17.12.2025", "TMC125307", "MUHAMMAD HASEEB", "MRI", 782),
        ("18.12.2025", "TMC125354", "SEYED KADER MOHAIDEEN SEYED", "MRI", 854),
        ("22.12.2025", "TMC125471", "NITIN KUMAR RAM PAL", "CT", 459),
        ("28.12.2025", "125725", "Mr TASHADDAK HOSSAIN", "MRI", 854),
        ("28.12.2025", "TMC125743", "MUSHTAQ MUHAMMAD WAZIR MUHAMMAD", "MRI", 854),
        ("31.12.2025", "TMC125914", "MOHAMMED OBAID", "MRI", 835,
         "Source sheet (Sep-Dec 2025 / Jan 2026) shows 31.12.2026; taken as 31-Dec-2025."),
    ],
    "Jan-2026": [
        ("02.01.2026", "TMC125957", "AYUB ALI", "MRI", 782),
        ("01.01.2026", "TMC070248", "FAHEED MUHAMMED", "MRI", 795),
    ],
    "Feb-2026": [
        ("01.02.2026", "127513", "Mr MUHAMMAD RIAZ", "MRI LUMBR C-", 782),
        ("02.02.2026", "127526", "Mr ZAHEER ABBAS GHULAM HUSSAIN", "MRI CRV C-", 795),
        ("02.02.2026", "127562", "Mr GIR BAHADUR DAMAI", "MRI LUMBR C-", 782),
        ("04.02.2026", "127667", "Mr NAQEEB ULLAH", "MRI LUMBR C-", 782),
        ("04.02.2026", "127661", "Mr MOHAMMED SHAYED", "MRI SHOULDER", 835),
        ("05.02.2026", "127680", "Mr AHMED ATTAULLA ALI AKBAR ATTAULLA", "MRI SHOULDER C-", 835),
        ("08.02.2026", "127796", "Mr SURENDRA YADAV BISUNDEYAL YADAV", "MRI LMBR C-", 782),
        ("08.02.2026", "127808", "Mr IBRAR AHMED KHAN", "MRI SHOULDER", 835),
        ("08.02.2026", "127810", "Mr SALMAN AHMAD FAIZAN AHMAD", "MRI LMBR C-", 782),
        ("09.02.2026", "127855", "Mr AURANGZEB QAMAR ZAMAN", "MRI LMBR C-", 782),
        ("10.02.2026", "127901", "Mr ZULFIQAR ALI", "MRI LUMBR C-", 782),
        ("12.02.2026", "128013", "Mr ABDUL ROUF", "MRI LUMBR C-", 782),
        ("12.02.2026", "4133", "Mr ARBAN ALI ABDUL RASHID", "MRI LUMBR C-", 782),
        ("12.02.2026", "128028", "Mr MUHAMMAD ASLAM", "MRI LUMBR C-", 782),
        ("12.02.2026", "128038", "Mr SOHIDUL ISLAM", "MRI SHOULDER", 835),
        ("13.02.2026", "128050", "UMER ZAD khan", "MRI KNEE C-", 854),
        ("14.02.2026", "128077", "Mr SADE KHAN RAZI KHAN", "MRI LMBR C-", 782),
        ("14.02.2026", "7366", "Mr BAHADAR ZAMAN", "MRI LUMBR C-", 782),
        ("16.02.2026", "128224", "Ms ANAGHA PANDARIKKAL", "MRI CRV C-", 795),
        ("17.02.2026", "128317", "Mr SHOYAIB ABUL KASHEM BIBI", "MRI LUMBR C-", 782),
        ("22.02.2026", "128499", "Mr MUHAMMED FAYYAS PUTHALATHTHAZHE", "MRI LMBR C-", 782),
        ("22.02.2026", "128508", "Mr SOMON BADSHA MIAH", "MRI KNEE C-", 854),
        ("22.02.2026", "128528", "Mr SHAHID ULLAH ABU SYED", "MRI R HAND", 891),
        ("24.02.2026", "128650", "Mr MOSLEH UDDIN", "MRI LUMBR C-", 782),
    ],
    "Mar-2026": [
        ("08.03.2026", "129054", "Mr UMAR HUSSAIN MIR BAZ", "MRI LMBR C-", 782),
        ("09.03.2026", "129096", "Mr ARSHAD KHAN SAR ZAMIN", "MRI LMBR C-", 782),
        ("10.03.2026", "129125", "Mr ALI AWAIS MUHAMMAD MEHBOOB", "MRI CRV C-", 795),
        ("12.03.2026", "129219", "Mr MOHAMMED NURUN NABI", "MRI CRV C-", 795),
        ("14.03.2026", "129247", "Mr MD MUBASHIR ALI YARIS ALI", "MRI FEMUR C-", 854),
        ("15.03.2026", "129285", "Mr MOHIN UDDIN", "MRI KNEE", 854),
        ("15.03.2026", "129286", "Mr EHAB FOUAD", "MRI LUMBR C-", 782),
        ("18.03.2026", "129403", "Mr MOHAMMED UDDIN", "MRI LUMBR C-", 782),
        ("22.03.2026", "129461", "Mr ARSALA KHAN", "MRI LUMBR C-", 782),
        ("24.03.2026", "129540", "Mr GANGADHAR ARRAM", "MRI LUMBR C-", 782),
        ("30.03.2026", "104157", "Mr YOUSUF MIAH", "MRI CRV C-", 795),
        ("30.03.2026", "129728", "Mr SHAFEEK PUTHUKKUDY", "MRI KNEE C-", 854),
        ("31.03.2026", "129756", "Mr UMAR YASEEN PIR BAD KHAN", "MRI LUMBR C-", 782),
    ],
    "Apr-2026": [
        ("01.04.2026", "128650", "Mr MOSLEH UDDIN AHMED", "MRI CRV C-", 795),
        ("02.04.2026", "129825", "Mr NAEEM SHAHZAD", "MRI LUMBR C-", 782),
        ("03.04.2026", "129852", "QASIM ALI BARKAT ALI", "MRI LEFT KNEE C-", 854),
        ("05.04.2026", "129927", "Mr ALI ASGHAR", "MRI LUMBR C-", 782),
        ("06.04.2026", "129965", "MUHAMMAD RAFIQ", "MRI LMBR C-", 782),
        ("11.04.2026", "130174", "Mr ZAFAR IQBAL MUHAMMAD", "MRI CRV C-", 795),
        ("15.04.2026", "130399", "Mr SAHIBZADA HIDAYAT", "MRI SHOULDER", 835),
        ("16.04.2026", "130458", "Mr SIYA RAM MURARI LAL", "MRI SHOULDER", 835),
        ("23.04.2026", "130777", "Mr HAJI REHMAN PIAO GUL", "MRI LUMBR C-", 782),
        ("27.04.2026", "130946", "Mr DINESH GOUD AMPALLY", "CT ABD & PLVS C-", 459),
        ("27.04.2026", "130994", "Mr MUHAMMAD IMRAN", "MRI LUMBR C-", 782),
        ("30.04.2026", "35346", "Mr MASHAD HALEEM HAYAT GUL", "MRI CRV C-", 795),
        ("30.04.2026", "54685", "Mr SAQIB SAEED MUHAMMAD SAEED", "MRI LT SHOULDER C-", 835),
        ("30.04.2026", "131116", "Mr SALIM KHAN AZIZ KHAN", "MRI SHOULDER", 835),
    ],
    "May-2026": [
        ("01.05.2026", "784-1979-5848498-8", "ANOWAR ABDUR ROB", "MRI", 854),
        ("05.05.2026", "784-1990-6593182-8", "ABID ULLAH MUNIR KHAN", "MRI", 854),
        ("05.05.2026", "784-1988-1479579-9", "VINOD VISWAMBHARAN VISWAMBHARAN NANU", "MRI", 782),
        ("12.05.2026", "784-1971-4871317-1", "RAJU BADDE BADDE DHARMA RAJU", "MRI", 795),
        ("08.05.2026", "784-1983-8290646-3", "ABDUL HAMID SHER JAN", "MRI", 782),
        ("21.05.2026", "784-1984-8507697-4", "ANWAR HOSSAIN KHAYER MIAH", "MRI", 746),
        ("10.05.2026", "784-1986-2851947-6", "MOHAMMED NURUN NABI NURUL AMIN", "MRI", 854),
        ("10.05.2026", "784-1977-9508054-4", "HABIBUR RAHAMAN NOOR MOHAMMED", "MRI", 782),
        ("11.05.2026", "784-1992-8031576-5", "IMRAN IQBAL KHAN", "MRI", 782),
        ("16.05.2026", "784-1988-8379697-5", "WASEEMULLAH FATEH KHAN", "MRI", 854),
        ("23.05.2026", "784-1989-2468094-9", "ALBANUS IFEANYI UZOMA", "MRI", 854),
        ("14.05.2026", "784-1991-5068348-2", "ARVIND KUMAR GUPTA RAM CHHABILA GUPTA", "MRI", 782),
        ("26.05.2026", "784-1992-0657616-9", "SONU KUMAR PARMESHVAR", "MRI", 782),
        ("23.05.2026", "784-1986-2854657-8", "FIROZ SALMANI", "MRI", 795),
        ("30.05.2026", "784-1982-6868161-0", "MOHAMMAD SALAH ALBNIA", "MRI", 782),
        ("30.05.2026", "784-1980-9063969-5", "YAR GUL JAN SAKHTA JAN", "MRI", 854),
    ],
    "Jun-2026": [
        ("01.06.2026", "784-1987-1405216-8", "MOHAMMED SHAYED HOSSEN SHAFIUR RAHMAN", "MRI", 782),
        ("06.06.2026", "784-1990-7249497-6", "MUHAMMAD MUSTAFA MUHAMMAD MANZOOR", "MRI", 835),
        ("10.06.2026", "784-2000-7203636-9", "USMAN ALI MUHAMMAD ASHRAF", "MRI", 795),
        ("25.06.2026", "784-1966-7194295-9", "Mr RAJU BALAKRISHNASAMY", "MRI", 854),
        ("10.06.2026", "784-1989-6832091-5", "EDRIS HAIDER KHEL", "MRI", 854),
        ("10.06.2026", "784-1987-8032685-0", "JANAT ULLAH MUHAMMAD YOUNAS", "MRI", 782),
        ("16.06.2026", "784-1994-4360482-4", "MOHAMMAD NUR KADER MOHAMMAD DANA MEAH", "MRI", 854),
        ("25.06.2026", "784-1983-5830687-0", "MANJUR ANSARI", "MRI", 871),
        ("21.06.2026", "784-1990-3624127-3", "YAR MUHAMMAD NIAZ MUHAMMAD", "MRI", 854),
        ("25.06.2026", "784-1992-1529043-0", "BARSHA SHRESTHA", "MRI", 782),
        ("28.06.2026", "784-1992-3107036-0", "NOOR UL HADI FAZALE RAHIM", "MRI", 782),
    ],
    # No radiology data received yet -> blank fill-in rows (Jul)
    "Jul-2026": [],
    # From Faiz_Patient_Data_August-2026 (no bill amounts in source)
    "Aug-2026": [
        ('03.08.2026', 'TVIP01017700', 'Faisal Khaled Awad Salem Alhanaee', 'MRI', None),
        ('05.08.2026', 'TVIP00383558', 'Noura Ramadan Mohammed Alblooshi', 'MRI', None),
        ('05.08.2026', 'TVIP01012803', 'Farseena Abdul', 'US PREGNANT', None),
        ('05.08.2026', 'TVIP00899382', 'Bushra Abdul', 'US PREGNANT', None),
        ('05.08.2026', 'TVIP01017948', 'Eslam Salahaldin', 'US PREGNANT', None),
        ('05.08.2026', 'TVIP01017956', 'Yassmin Wael', 'US PREGNANT', None),
        ('05.08.2026', 'TVIP01012463', 'Mai Youssef', 'US PREGNANT', None),
        ('06.08.2026', 'TVIP00870626', 'Aysha Saeed Abdulla Yaser Alkaabi', 'MRI', None),
        ('06.08.2026', 'TVIP01018050', 'Theertha Puthiyapuray', 'US PREGNANT', None),
        ('06.08.2026', 'TVIP00898744', 'Nihala Sherin Kaipadan Chalil', 'US PREGNANT', None, "Moved from Aug visits: the August-2026 file lists this as a radiology referral (name there 'Nihala Sherin'; full name from the marketing file)."),
        ('06.08.2026', 'TVIP00922585', 'Beverly Ngajo Albaracin', 'US PREGNANT', None),
        ('08.08.2026', '784-1985-3403496-6', 'Jay Singh Bohara', 'MRI', None),
        ('08.08.2026', 'TVIP01018290', 'Niha Fathima', 'US PREGNANT', None),
        ('08.08.2026', 'TVIP01018291', 'Noufiya Salahudeen', 'US PREGNANT', None),
        ('09.08.2026', 'TVIP00063874', 'Amal Abdallah Alnuaimi', 'MRI', None),
        ('09.08.2026', '526008408', 'Mai Suhail Mohammed Al Muharrami', 'MRI', None),
        ('10.08.2026', '784-1986-5350376-7', 'Bipon Kumar Barua Sunil Barua', 'MRI', None),
        ('10.08.2026', '784-1980-4036273-7', 'Fazal Jan Zahir Jan', 'CT', None),
        ('12.08.2026', 'TRM027700', 'Shanda Abdulla Khamis', 'MRI', None),
        ('12.08.2026', 'TVIP01018828', 'Sadia Sultana', 'US PREGNANT', None),
        ('12.08.2026', 'TVIP01015684', 'Sadia Bibi', 'US PREGNANT', None),
        ('13.08.2026', '506437977', 'Rubayea Salem Rubayea Aldhaheri', 'MRI', None),
        ('13.08.2026', 'TVIP01018885', 'Semeena Varambanalu', 'US PREGNANT', None),
        ('13.08.2026', 'TVIP01018920', 'Malaz Elbager', 'US PREGNANT', None),
        ('16.08.2026', 'TVIP01019170', 'Shamla Kallarakkal', 'US PREGNANT', None),
        ('16.08.2026', 'TVIP01019178', 'Muhsina Thattarathod', 'US PREGNANT', None),
        ('16.08.2026', 'TVIP01012813', 'Noor Bina', 'US PREGNANT', None),
        ('16.08.2026', 'TVIP01019155', 'Tanjina Akther', 'US TRANSVAGINAL', None),
        ('17.08.2026', '784-2006-3523149-9', 'Sonu Jagga Singh', 'CT', None),
        ('17.08.2026', 'TVIP01019287', 'Aisha Sarwar', 'US PREGNANT', None),
        ('18.08.2026', '566707100', 'Aisha Mohamed Abdulghani', '3-MRI', None, 'Source gives file 566707100 for both Aisha Mohamed Abdulghani and Aidah Eisa Saleh Alruqaishi on 18 Aug.'),
        ('18.08.2026', '566707100', 'Aidah Eisa Saleh Alruqaishi', '3-MRI', None, 'Source gives file 566707100 for both Aisha Mohamed Abdulghani and Aidah Eisa Saleh Alruqaishi on 18 Aug.'),
        ('19.08.2026', '506194419', 'Abdulla Ahmed Ibrahim Alobeidli', 'MRI', None),
        ('19.08.2026', 'TVIP01019555', 'Simi Aruldhas', 'US PREGNANT', None),
        ('19.08.2026', 'TVIP01013011', 'Maryam Imtiaz', 'US PREGNANT', None),
        ('22.08.2026', '503841100', 'Saif Obaid Saif Alwaly Alkindi', 'MRI', None),
        ('23.08.2026', '509322423', 'Khuloud Taleb Mohammed Almemari', 'MRI', None),
        ('23.08.2026', 'TVIP01019927', 'Irudaya Ancy', 'US PREGNANT', None),
        ('24.08.2026', '507631090', 'Sabha Ali Obaid Ali Alazeezi', 'MRI', None),
        ('26.08.2026', 'TVIP01020441', 'Israt Jahan', 'US PREGNANT', None),
        ('26.08.2026', 'TVIP01020489', 'Ebtissam Abdelrahmin', 'US PREGNANT', None),
        ('27.08.2026', '504445500', 'Sultan Amhi Humaid Amhi Almansoori', 'MRI', None),
        ('29.08.2026', '784-1996-2778384-0', 'Muhammad Tahir Rab Nawaz', 'MRI', None),
        ('29.08.2026', '507831138', 'Mariam Hamad Rashed Alnuaimi', 'MRI', None),
        ('30.08.2026', '784-1986-5915314-6', 'Usankutty Na Tukallu', 'MRI', None),
        ('30.08.2026', '507532203', 'Fatima Saeed Abdulrahman Alkaabi', 'MRI', None),
        ('30.08.2026', 'TVIP01020965', 'Wesam Ahmed', 'US PREGNANT', None),
        ('30.08.2026', 'TVIP01020834', 'Ajna Nizar Nizar S J', 'US PREGNANT', None),
    ],
    # From Faiz_Patient_Data_September-2026 (no bill amounts in source)
    "Sep-2026": [
        ('01.09.2026', 'TVIP01021143', 'Shabbir Ahmad Nasar Ul Din', 'CT SCAN', None),
        ('02.09.2026', 'TVIP01014023', 'Supriya Karapu Laxminarayana Puli', 'US PREGNANT', None),
        ('02.09.2026', 'TVIP01021249', 'Sumaira Naz Usman Rashid', 'US PREGNANT', None),
        ('02.09.2026', 'TVIP01015398', 'Ameena Padinjarevattatheput Hukkas', 'US PREGNANT', None),
        ('03.09.2026', 'TVIP01021417', 'Shahala Mol Vellarama', 'US PREGNANT', None),
        ('03.09.2026', 'TVIP01013116', 'Fathima Risvana Madathil Abdul Rasak', 'US PREGNANT', None),
        ('03.09.2026', 'TVIP01014947', 'Mercy Kpeta', 'US PREGNANT', None),
        ('04.09.2026', '1013143', 'Jasna Asharaf Kunhayi', 'US PREGNANT', None),
        ('05.09.2026', 'TVIP01021581', 'Reem Ghasan Morshed', 'US PREGNANT', None),
        ('06.09.2026', 'TVIP00854813', 'Laila Rehman Abdul Rehman', 'US PREGNANT', None),
        ('06.09.2026', 'TVIP01021758', 'Shaikha Ali Thani Ali Alshamsi', 'MRI CRV', None),
        ('07.09.2026', 'TVIP01012534', 'Nadal Abdelelah Siddig Elzain', 'US PREGNANT', None),
        ('08.09.2026', 'TVIP01021984', 'Sobia Rashid Muhammad Rashid', 'US PREGNANT', None),
        ('09.09.2026', 'TVIP01022197', 'Wajid Ali Piao Gul', 'MRI LUMBR', None),
        ('12.09.2026', 'TVIP01022461', 'Ameera Haroon Haroon', 'US PREGNANT', None),
        ('12.09.2026', 'TVIP01022494', 'Musanif Shah Abdullah Shah', 'MRI ARM', None),
        ('12.09.2026', 'TVIP00402597', 'Fatima Salama Alkatherei', 'MRI CRV', None),
        ('13.09.2026', 'TVIP01022562', 'Sumaiya Mohammed Yasir Hamza', 'US PREGNANT', None),
        ('13.09.2026', 'TVIP01013012', 'Asniya Valiyaveedu Abdul Kareem Abdul', 'US PREGNANT', None),
        ('13.09.2026', 'TVIP01022522', 'Haider Khan Muhammad Gul', 'MRI LMBR', None),
        ('14.09.2026', 'TVIP01022788', 'Khizar Hayat Qamar Abbas', 'CT SCAN', None),
        ('14.09.2026', 'TVIP00983134', 'Enayat Ullah Gulab Din Marwat', 'MRI LUMBR', None),
        ('14.09.2026', 'TVIP00965236', 'Shahid Zainudheen Zainudheen Kakkadavan Parammal', 'CT MAX', None),
        ('15.09.2026', 'TVIP01013619', 'Uzma Mustaqeem Mohd Mustaqeem', 'US PREGNANT', None),
        ('15.09.2026', 'TVIP01022844', 'Muhammad Asif Mir Kazim Mehmood', 'CT MAX', None),
        ('15.09.2026', 'TVIP01022861', 'Nasser Mohammed Khalfan Al Neyadi', 'MRI B SHOULDER', None),
        ('16.09.2026', 'TVIP01023036', 'Ranjana Majhi', 'US PREGNANT', None),
        ('16.09.2026', 'TVIP01023039', 'Shamima Haque Nishi', 'US PREGNANT', None),
        ('17.09.2026', 'TVIP01014624', 'Ahsanah Rahman Azizur Rahman', 'US PREGNANT', None),
        ('17.09.2026', 'TVIP01023249', 'Anusha Sandamali Maltunga Arachchile', 'MRI LUMBR', None),
        ('17.09.2026', 'TVIP01023104', 'Amal Abdulla Saad Alshanini', 'MRI KNEE', None),
        ('19.09.2026', 'TVIP01023366', 'Fatema Israt Mohammad Alamgir Alam', 'US PREGNANT', None),
        ('19.09.2026', 'TVIP01023324', 'Ariane Kaye Bucasas Badua', 'US PREGNANT', None),
        ('19.09.2026', 'TVIP01023351', 'Abhishek Gond Narvadeshwar Gond', 'CT SCAN', None),
        ('22.09.2026', 'TVIP01023846', 'Osman Rohani', 'CT SCAN', None),
        ('22.09.2026', 'TVIP01023860', 'Safeer Kunnath Moiou Olavan Kunnath', 'MRI LUMBR', None),
        ('22.09.2026', 'TVIP00851673', 'Muhammad Waseem Ashraf Muhammad', 'MRI LUMBR', None),
        ('23.09.2026', 'TVIP01023924', 'Uzma Siraj Siraj Jamil', 'US PREGNANT', None),
        ('26.09.2026', 'TVIP01024200', 'Aliza Anum Syed Mohammed Zaid Khawja', 'US PREGNANT', None),
        ('26.09.2026', 'TVIP01024211', 'Kala Miah Abo Ddor Ali Ddor', 'CT SCAN', None),
        ('26.09.2026', 'TVIP00970990', 'Alia Huwaiden Thiban Ali Alketbi', 'MRI KNEE', None),
        ('27.09.2026', 'TVIP01023351', 'Abhishek Gond Narvadeshwar Gond', 'MRI HIP WC', None),
        ('29.09.2026', 'TVIP00579183', 'Emaan Blaisha Muhammed', 'MRI', None),
        ('30.09.2026', 'TVIP01024887', 'Tilak Bahadur', 'MRI', None),
        ('30.09.2026', 'TVIP01024882', None, 'US RETRO', None, 'Patient name not given in the source file (Faiz_Patient_Data_September-2026).'),
        ('30.09.2026', 'TVIP01023036', 'Ranjana Majhi', 'US PREGNANT', None),
    ],
}
BLANK_ROWS = 5


def clean_name(n):
    if not n:
        return None
    parts = n.split()
    if parts[0].lower() in ("mr", "ms", "mrs"):
        parts = parts[1:]
    return " ".join(p.capitalize() for p in parts)


def put(ws, ref, src_cell, value=None):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font, c.fill, c.border = copy(src_cell.font), copy(src_cell.fill), copy(src_cell.border)
    c.alignment, c.number_format = copy(src_cell.alignment), src_cell.number_format
    return c


def main(path):
    wb = openpyxl.load_workbook(path)
    T = wb["Oct-2026"]  # style template
    # locate the template's radiology total rows by label (rows can shift when Oct gets new entries)
    TC = next(r for r in range(1, T.max_row + 1) if T.cell(r, 4).value == "Total Referrals Count")
    TB = TC + 1
    for sheet, rows in RADIOLOGY.items():
        ws = wb[sheet]
        t = ws.max_row          # existing TOTAL PATIENTS row
        s = t - 1               # teal spacer row
        last = s - 1            # last main-table data row
        month = ws["A2"].value.replace("PATIENT DATA OF ", "")

        # -- widen the main table to the Oct model (A:F, Bill Amount column) --
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
        for r in (1, 2):
            put(ws, f"F{r}", T[f"F{r}"])
        ws.merge_cells("A1:F1"); ws.merge_cells("A2:F2")
        put(ws, "F3", T["F3"], "Bill Amount (AED)")
        for r in range(4, s):
            put(ws, f"F{r}", T["F4"])
        for col in "ABCDEF":
            put(ws, f"{col}{s}", T[f"{col}49"])
        ws.merge_cells(f"A{s}:F{s}")
        for col in "ABCDEF":
            put(ws, f"{col}{t}", T[f"{col}50"])
        ws.merge_cells(f"A{t}:D{t}")
        ws.column_dimensions["F"].width = T.column_dimensions["F"].width

        b = t + 1  # TOTAL BILL AMOUNT row
        for col in "ABCDEF":
            put(ws, f"{col}{b}", T[f"{col}51"])
        ws[f"A{b}"] = f"TOTAL BILL AMOUNT (AED) - {month}"
        ws[f"F{b}"] = f"=SUM(F4:F{last})"
        ws.merge_cells(f"A{b}:E{b}")
        ws.row_dimensions[b].height = T.row_dimensions[51].height

        # -- radiology section --
        cp, cb = b + 2, b + 3                 # combined rows
        title, hdr = cb + 2, cb + 3
        data = sorted(rows, key=lambda x: datetime.strptime(x[0], "%d.%m.%Y"))
        n = len(data) or BLANK_ROWS
        first, lastr = hdr + 1, hdr + n
        cnt, bill = lastr + 2, lastr + 3

        for col in "ABCDEF":
            put(ws, f"{col}{title}", T[f"{col}56"])
            put(ws, f"{col}{hdr}", T[f"{col}57"], T[f"{col}57"].value)
        ws[f"A{title}"] = "RADIOLOGY DATA - OUT PATIENTS"
        ws.merge_cells(f"A{title}:F{title}")
        ws.row_dimensions[title].height = T.row_dimensions[56].height
        ws.row_dimensions[hdr].height = T.row_dimensions[57].height

        for i in range(n):
            r = first + i
            for col in "ABCDEF":
                put(ws, f"{col}{r}", T[f"{col}58"])
            ws.row_dimensions[r].height = T.row_dimensions[58].height
            if not data:
                continue
            d, fno, name, exam, amt, *note = data[i]
            ws[f"A{r}"] = i + 1
            ws[f"B{r}"] = datetime.strptime(d, "%d.%m.%Y")
            ws[f"C{r}"] = fno
            ws[f"D{r}"] = clean_name(name)
            ws[f"E{r}"] = exam
            ws[f"F{r}"] = amt
            if note:
                ws[f"B{r}" if "31.12" in d else f"D{r}"].comment = Comment(note[0], "Claude")

        put(ws, f"D{cnt}", T[f"D{TC}"], "Total Referrals Count")
        put(ws, f"E{cnt}", T[f"E{TC}"], f"=COUNTA(C{first}:C{lastr})")
        put(ws, f"D{bill}", T[f"D{TB}"], "Total Bill Amount")
        put(ws, f"E{bill}", T[f"E{TB}"])
        put(ws, f"F{bill}", T[f"F{TB}"], f"=SUM(F{first}:F{lastr})")
        ws.merge_cells(f"D{bill}:E{bill}")
        for r in (cnt, bill):
            ws.row_dimensions[r].height = T.row_dimensions[TC].height

        for col in "ABCDEF":
            put(ws, f"{col}{cp}", T[f"{col}53"])
            put(ws, f"{col}{cb}", T[f"{col}54"])
        ws[f"A{cp}"] = "COMBINED TOTAL PATIENTS / REFERRALS"
        ws[f"E{cp}"] = f"=E{t}+E{cnt}"
        ws[f"A{cb}"] = "COMBINED TOTAL BILL AMOUNT (AED)"
        ws[f"F{cb}"] = f"=F{b}+F{bill}"
        ws.merge_cells(f"A{cp}:D{cp}"); ws.merge_cells(f"A{cb}:E{cb}")
        for r in (cp, cb):
            ws.row_dimensions[r].height = T.row_dimensions[53].height
        print(sheet, "referrals", len(data), "rows", t, "->", bill)

    # -- Summary: add radiology columns --
    S = wb["Summary"]
    for rng in list(S.merged_cells.ranges):
        S.unmerge_cells(str(rng))
    for r in (1, 2):
        put(S, f"D{r}", S[f"C{r}"]); put(S, f"E{r}", S[f"C{r}"])
    put(S, "D3", S["C3"], "Radiology Referrals")
    put(S, "E3", S["C3"], "Radiology Bill (AED)")
    for i, sheet in enumerate(wb.sheetnames[1:]):
        r = 4 + i
        ws = wb[sheet]
        cnt_row = next(x for x in range(ws.max_row, 3, -1) if ws[f"D{x}"].value == "Total Referrals Count")
        put(S, f"D{r}", S[f"C{r}"], f"='{sheet}'!E{cnt_row}")
        put(S, f"E{r}", S[f"C{r}"], f"='{sheet}'!F{cnt_row + 1}").number_format = "#,##0.00"
    put(S, "D16", S["C16"]); put(S, "E16", S["C16"])
    put(S, "D17", S["C17"], "=SUM(D4:D15)")
    put(S, "E17", S["C17"], "=SUM(E4:E15)").number_format = "#,##0.00"
    for ref in ("A1:E1", "A2:E2", "A16:E16", "A17:B17", "A19:E19"):
        S.merge_cells(ref)
    S.column_dimensions["D"].width = 22
    S.column_dimensions["E"].width = 22
    S["A19"] = (S["A19"].value.rstrip(".") + ". Radiology referrals: from the Spectrum Center "
                "referral sheets / radiology screenshots; Aug & Sep 2026 from the August/September-2026 files; none received yet for Jul 2026.")
    S.row_dimensions[19].height = 45

    wb.save(path)


if __name__ == "__main__":
    main(sys.argv[1])
