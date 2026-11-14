"""Fill bill amounts (column F) from the Nov-2026 attachments.

Run last, after scripts/add_radiology_sections.py:
    python scripts/add_bill_amounts_nov2026.py <master.xlsx>

Rows are located by sheet section + date + file number + patient name (never by row number).
Missing amounts stay blank until the data is cleared (repo rule).
"""
import sys
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

# July visits - faz1_1.xlsx, sheet "JULY 2026" (amounts per doctor are added together).
# (date, file, name, total AED, note)
JUL_VISITS = [
    ('2026-07-01', 'TVIP01013249', 'Ehsan Ullah', 459.0, None),
    ('2026-07-01', 'TVIP00980784', 'Sumithra Venkatesan', 153.36, None),
    ('2026-07-02', 'TVIP01013128', 'Muhammed Ahyaan', 295.0, None),
    ('2026-07-04', 'TVIP01013800', 'Mansoor Ali Korankulangara', 394.0, None),
    ('2026-07-05', 'TVIP01013128', 'Muhammed Ahyaan', 128.0, None),
    ('2026-07-06', 'TVIP00829926', 'Leela Rairoth', 529.2, 'Bill breakdown: AFSHEEN,IHAB: 454.2/75'),
    ('2026-07-13', 'TVIP01015079', 'Abdul Gafoor Kummalil', 75.0, None),
    ('2026-07-14', 'TVIP01015139', 'Eassa Karuvante', 85.0, None),
    ('2026-07-14', 'TVIP00830036', 'Bushra Bilal Muhammad', 250.0, None),
    ('2026-07-16', 'TVIP00902268', 'Muhammed Ali Ponneth', 556.0, None),
    ('2026-07-18', 'TVIP00658599', 'Maheen Abdulrahuman', 619.4, None),
    ('2026-07-18', 'TVIP00902268', 'Muhammed Ali Ponneth', 589.0, None),
    ('2026-07-19', 'TVIP00961592', 'Muhammad Zaroob', 176.6, None),
    ('2026-07-19', 'TVIP00654762', 'Shahzad Afsar', 489.6, 'Bill breakdown: AFSHEEN,NANDINI: 75 /414.60'),
    ('2026-07-20', 'TVIP01015970', 'Mohammed Sabbir Ahmed', 253.4, None),
    ('2026-07-20', 'TVIP01015196', 'Adel Fares Afifi Ibrahim Afifi', 430.0, 'Bill breakdown: MOFEED,  IHAB , ALAMGIR: 75 / 85 / 270'),
    ('2026-07-20', 'TVIP01015969', 'Rubel Miah Suruj Miah', 236.6, None),
    ('2026-07-21', 'TVIP00979556', 'Zayan Hamad Maheen', 289.2, 'Bill breakdown: ROAA,NANDINI: 85 / 204.2'),
    ('2026-07-21', 'TVIP00658599', 'Maheen Abdulrahuman', 346.0, 'Bill breakdown: IMAD,NANDINI: 346 /0'),
    ('2026-07-21', 'TVIP01016042', 'Shihabudheen Palakkal', 52.6, None),
    ('2026-07-25', 'TVIP00973129', 'Hiba Kunnummal', 136.4, None),
    ('2026-07-25', 'TVIP01015970', 'Mohammed Sabbir Ahmed', 253.4, None),
    ('2026-07-25', 'TVIP01015969', 'Rubel Miah Suruj Miah', 236.6, None),
    ('2026-07-25', 'TVIP00902268', 'Muhammed Ali Ponneth', 85.0, 'Bill breakdown: FERAS, ALI: 0 / 85'),
    ('2026-07-25', 'TVIP01013128', 'Muhammed Ahyaan', 374.0, None),
    ('2026-07-30', 'TVIP01017072', 'Chandran Velath', 316.0, 'Bill breakdown: MUSTAFA,RADI,HIGAZY , TAHA: 231 /0/85/0'),
    ('2026-07-30', 'TVIP01017073', 'Ramakrishnan Melekadungoth', 257.0, 'Bill breakdown: MUSTAFA,HIGAZY,SHIRAZ: 162/95/0'),
    ('2026-07-30', 'TVIP01017075', 'Hammad Hussain Janjua', 505.69, 'Bill breakdown: MUSTAFA,ABDULHALEEM: 485 / 20.69'),
    ('2026-07-30', 'TVIP00865589', 'Nijam Uddin Mubarok Ali', 884.2, 'Bill breakdown: MOZAMMIL,SIBTAIN,RADI: 194.20 /493.00 /197.00'),
    ('2026-07-30', 'TVIP01017118', 'Rodello Olazo Santos', 731.1, None),
    ('2026-07-30', 'TVIP01017120', 'Brudi Solamon Korattikunnan', 731.1, None),
    ('2026-07-30', 'TVIP01017122', 'Ibrahim Thayyil Hamza', 104.5, None),
    ('2026-07-30', 'TVIP01017126', 'Jeganathan Perumal', 851.0, None),
    ('2026-07-30', 'TVIP01017132', 'Haneefa Thekkumbath', 877.0, None),
    ('2026-07-30', 'TVIP01017134', 'Mohammad Asim Nasimur Rehaman', 857.0, None),
    ('2026-07-30', 'TVIP01017151', 'Balu Chandrashekar', 1184.0, 'Bill breakdown: ROAA,ABDULHALEEM,ALI,JIMSHAD: 95/404/415/270'),
    ('2026-07-30', 'TVIP01001897', 'Mohammed Ameer Puzhangaraillath', 425.2, None),
    ('2026-07-30', 'TVIP01017076', 'Jamghat Gul', 160.0, 'Bill breakdown: MUSTAFA,ALAA: 75/85'),
    ('2026-07-30', 'TVIP01017168', 'Jishnu Thamarassery', 251.65, 'Bill breakdown: ROAA,ALI: 127 /124.65'),
    ('2026-07-30', 'TVIP01017291', 'Hood Yaliq Firos', 45.0, None),
    ('2026-07-30', 'TVIP01017171', 'Unnikrishnan Thelasseri Parambil', 350.0, 'Bill breakdown: ROAA,TAGREED: 350/0'),
    ('2026-07-30', 'TVIP01017173', 'Vijayakumar Arumugam', 4712.0, 'Bill breakdown: ROAA,AGNEL,ALI: 390/817/3505'),
    ('2026-07-30', 'TVIP01017170', 'Renjith Raju', 75.0, 'Bill breakdown: ROAA,TAGREED: 0/75'),
    ('2026-07-30', 'TVIP01017184', 'Junaid Ahmed Jamaluddin', 601.0, 'Bill breakdown: MOZAMMIL,AGNEL: 220/381'),
    ('2026-07-30', 'TVIP01017181', 'Naze Ahmed Rashed', 430.0, None),
    ('2026-07-30', 'TVIP01017186', 'Sahid Mahammad Nur Mahammad', 893.0, None),
    ('2026-07-30', 'TVIP01017198', 'Anas Makkikaran', 398.0, None),
    ('2026-07-30', 'TVIP01017289', 'Firos Cheriya', 45.0, None),
    ('2026-07-30', 'TVIP01017206', 'Rana Muhammad Zubair', 518.6, 'Bill breakdown: DIMEH,AGNEL: 466/52.6'),
    ('2026-07-30', 'TVIP01017231', 'Mohamed Fazil Erambath Alassan', 894.0, None),
    ('2026-07-30', 'TVIP01017232', 'Riyas Pilakkot Mohamed Kutty', 1443.0, 'Bill breakdown: MOZAMMIL,ALAA: 894 / 549'),
    ('2026-07-30', 'TVIP01017188', 'Abdul Qayum Zar Ur Rehman', 2525.0, 'Bill breakdown: ROAA,ZOUBI: 621 / 1904'),
    ('2026-07-30', 'TVIP01017250', 'Mohammedkutty Edathadathil', 415.0, None),
    ('2026-07-30', 'TVIP01017253', 'Abdul Rahim', 1002.0, 'Bill breakdown: ROAA,SIBTAIN: 496 / 506'),
    ('2026-07-30', 'TVIP01017260', 'Soukath Ali Kuthubuddin', 742.0, 'Bill breakdown: ROAA,BASSIMA: 114 / 628'),
    ('2026-07-30', 'TVIP00974777', 'Saber Rezk Elhussein', 386.0, None),
    ('2026-07-30', 'TVIP01017279', 'Mohamed Saber Rezk Elhusseini', 415.0, None),
    ('2026-07-30', 'TVIP01017276', 'Awatef Ramadan Hussein', 462.0, None),
    ('2026-07-30', 'TVIP01017284', 'Thomas Mathew', 327.6, 'Bill breakdown: ROAA,AGNEL: 140 / 187.6'),
    ('2026-07-30', 'TVIP01017285', 'Ajosh David Suresh', 671.0, None),
    ('2026-07-30', 'TVIP01017286', 'Swaraj Kurup Chandrasekharan', 438.6, None),
    ('2026-07-30', 'TVIP01017287', 'Sultan Ahmed Nizam Uddin', 386.0, None),
    ('2026-07-30', 'TVIP01017288', 'Abdul Rahiman Badiadka', 398.0, None),
    ('2026-07-30', 'TVIP01017290', 'Eiwa Yaliq', 75.0, None),
    ('2026-07-30', 'TVIP00963637', 'Nahala Variammadathil', 45.0, None),
    ('2026-07-30', 'TVIP01017293', 'Ameer Arinjikkal', 946.0, None),
    ('2026-07-30', 'TVIP00615797', 'Abdurahman Olimeah', 946.0, None),
    ('2026-07-30', 'TVIP01017292', 'Mohammed Harunur Rashid Delwar', 946.0, None),
    ('2026-07-30', 'TVIP00948643', 'Muhammed Hameed', 398.0, None),
    ('2026-07-30', 'TVIP01017204', 'Wafa Bint Mohammed Ameer', 427.9, None),
    ('2026-07-30', 'TVIP01017200', 'Nuha Bint Mohammed Ameer', 591.2, None),
    ('2026-07-30', 'TVIP00925005', 'Sanitha Valiyakath', 549.7, None),
    ('2026-07-30', 'TVIP01017202', 'Huda Bint Mohammed Ameer', 476.5, None),
    ('2026-07-30', 'TVIP01017234', 'Mohammed Shahid Ullah Mohammed', 1324.2, None),
    ('2026-07-31', 'TVIP01017311', 'Pramod Radhakrishna Pillai', 398.0, None),
    ('2026-07-31', 'TVIP01017312', 'Sureshkumar Chandrasekhara Pillai', 75.0, None),
    ('2026-07-31', 'TVIP01017335', 'Turki Qasem', 160.0, 'Bill breakdown: MUSTAFA: 75; ZAFER: 85'),
    ('2026-07-31', 'TVIP01017326', 'Muhammed Shafi Thazhathodi', 399.0, None),
    ('2026-07-31', 'TVIP01017325', 'Mohammed Shafeeq Thazhathodiyil', 1012.0, None),
    ('2026-07-31', 'TVIP01017297', 'Vinoth Kumar Murugan', 365.0, None),
    ('2026-07-31', 'TVIP01017307', 'Rusna Muchamad Rambang', 495.8, None),
    ('2026-07-31', 'TVIP01017324', 'Mohammed Hadi', 399.0, None),
    ('2026-07-31', 'TVIP01017118', 'Rodello Olazo Santos', 64.0, 'Bill breakdown: NIGHET: 0; AGNEL: 64'),
    ('2026-07-31', 'TVIP01017334', 'Marwa Ameen', 398.0, None),
    ('2026-07-31', 'TVIP01017333', 'Abdul Basith Manalodi', 485.0, None),
    ('2026-07-31', 'TVIP00604064', 'Sabil Ullah Hajat Khan', 849.0, None),
    ('2026-07-31', 'TVIP01017331', 'Hamid Mohammed Qasem Ahmed', 398.0, None),
]

# August - photos of the printed Aug visit list (rows 1-52, handwritten total 22195).
# (section, date, file, name, AED, note)
AUG = [
    ('visit', '2026-08-01', 'TVIP009213588', 'Tawoor Khan', 348, None),
    ('visit', '2026-08-01', 'TVIP01017126', 'Jeganathan Perumal', 315, None),
    ('visit', '2026-08-01', 'TVIP01017333', 'Abdul Basith Manalodi', 355, None),
    ('visit', '2026-08-01', 'TVIP01017200', 'Nuha Bint Mohammed Ameer', 0, None),
    ('visit', '2026-08-01', 'TVIP01017204', 'Wafa Bint Mohammed Ameer', 0, None),
    ('visit', '2026-08-01', 'TVIP01017443', 'Mohammed Rafeekh Thottungal', 334, None),
    ('visit', '2026-08-01', 'TVIP01017441', 'Mohamed Ali Thazhathody', 399, None),
    ('visit', '2026-08-01', 'TVIP01017403', 'Deepak Sahadevan Sumangala', 1837, None),
    ('visit', '2026-08-01', 'TVIP01017462', 'Mohammed Ashique Chettiyamthodi', 399, None),
    ('visit', '2026-08-01', 'TVIP01017438', 'Sakkeer Babu Chemmala', 399, None),
    ('visit', '2026-08-01', 'TVIP01017456', 'Sidheek Ayappally Saidhalavi', 1012, None),
    ('visit', '2026-08-01', 'TVIP01017455', 'Mohamed Afsal Vettan Shahul Hameed', 1012, 'Photo lists this visit twice (rows 2 and 13: same patient, day, doctor Roaa, AED 1012) - counted once.'),
    ('visit', '2026-08-02', 'TVIP00975328', 'Muhammed Naeem', 458.2, None),
    ('visit', '2026-08-02', 'TVIP01017490', 'Maysa', 492, None),
    ('visit', '2026-08-02', 'TVIP01017181', 'Naze Ahmed Rashed', 0, None),
    ('visit', '2026-08-03', 'TVIP00975328', 'Muhammed Naeem', 498, None),
    ('visit', '2026-08-03', 'TVIP01017627', 'Krishna Priya', 2327, None),
    ('visit', '2026-08-03', 'TVIP01017624', 'Joseph', 742.8, None),
    ('visit', '2026-08-03', 'TVIP00830448', 'Kavitha Satheesh', 2889.2, None),
    ('visit', '2026-08-03', 'TVIP01017120', 'Brudi Solamon Korattikunnan', 0, None),
    ('visit', '2026-08-03', 'TVIP01017122', 'Ibrahim Thayyil Hamza', 0, None),
    ('visit', '2026-08-04', 'TVIP00975328', 'Muhammed Naeem', 0, None),
    ('visit', '2026-08-04', 'TVIP01017627', 'Krishna Priya', 655, None),
    ('visit', '2026-08-04', 'TVIP01001897', 'Mohammed Ameer Puzhangaraillath', 0, None),
    ('visit', '2026-08-04', 'TVIP01017331', 'Hamid Mohammed Qasem Ahmed', 23.46, None),
    ('visit', '2026-08-04', 'TVIP00604064', 'Sabil Ullah Hajat Khan', 95, None),
    ('visit', '2026-08-04', 'TVIP01017334', 'Marwa Ameen', 0, None),
    ('visit', '2026-08-04', 'TVIP01017232', 'Riyas Pilakkot Mohamed Kutty', 549, None),
    ('visit', '2026-08-04', 'TVIP01017231', 'Mohamed Fazil Erambath Alassan', 0, None),
    ('visit', '2026-08-04', 'TVIP01017456', 'Sidheek Ayappally Saidhalavi', 0, None),
    ('visit', '2026-08-04', 'TVIP01017455', 'Mohamed Afsal Vettan Shahul Hameed', 85, None),
    ('visit', '2026-08-05', 'TVIP01015196', 'Adel Fares Afifi Ibrahim Afifi', 85, None),
    ('visit', '2026-08-06', 'TVIP01017126', 'Jeganathan Perumal', 0, None),
    ('visit', '2026-08-06', 'TVIP01018172', 'Majida Farsana', 316.8, None),
    ('visit', '2026-08-06', 'TVIP01017206', 'Rana Muhammad Zubair', 270, None),
    ('visit', '2026-08-06', 'TVIP01017333', 'Abdul Basith Manalodi', 901, None),
    ('rad', '2026-08-06', 'TVIP00898744', 'Nihala Sherin Kaipadan Chalil', 350, None),
    ('visit', '2026-08-07', 'TVIP01017151', 'Balu Chandrashekar', 835, None),
    ('visit', '2026-08-07', 'TVIP01017232', 'Riyas Pilakkot Mohamed Kutty', 0, None),
    ('visit', '2026-08-07', 'TVIP01017326', 'Muhammed Shafi Thazhathodi', 0, None),
    ('visit', '2026-08-07', 'TVIP01017443', 'Mohammed Rafeekh Thottungal', 0, None),
    ('visit', '2026-08-07', 'TVIP01017441', 'Mohamed Ali Thazhathody', 0, None),
    ('visit', '2026-08-07', 'TVIP01017438', 'Sakkeer Babu Chemmala', 288, None),
    ('visit', '2026-08-07', 'TVIP01018216', 'Sakkeer Puthu Parambil Pocker', 399, None),
    ('visit', '2026-08-08', 'TVIP01017462', 'Mohammed Ashique Chettiyamthodi', 0, None),
    ('visit', '2026-08-08', 'TVIP01017325', 'Mohammed Shafeeq', 75, None),
    ('visit', '2026-08-09', 'TVIP01018172', 'Majida Farsana', 414, None),
    ('visit', '2026-08-09', 'TVIP01017455', 'Mohamed Afsal Vettan Shahul Hameed', 75.2, None),
    ('visit', '2026-08-11', 'TVIP01017311', 'Pramod Radhakrishna Pillai', 114, None),
    ('visit', '2026-08-11', 'TVIP01017333', 'Abdul Basith Manalodi', 169, None),
    ('visit', '2026-08-05', 'TVIP00796985', 'Ahmed Abou Elfetouh Helal Elesawi', 1667, None),
]


def find(ws, section, d, fno, name):
    dd = datetime.strptime(d, "%Y-%m-%d")
    sec, hits = "visit", []
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 1).value == "RADIOLOGY DATA - OUT PATIENTS":
            sec = "rad"
        if sec == section and ws.cell(r, 2).value == dd and str(ws.cell(r, 3).value).replace(" ", "") == fno \
                and ws.cell(r, 4).value == name:
            hits.append(r)
    assert len(hits) == 1, (ws.title, section, d, fno, name, hits)
    return hits[0]


def put(ws, r, amt, note):
    c = ws.cell(r, 6)
    assert c.value in (None, amt), (ws.title, r, c.value, amt)
    c.value = amt
    c.number_format = "#,##0.00"
    if note:
        c.comment = Comment(note, "Claude")


def main(path):
    wb = openpyxl.load_workbook(path)
    for d, f, n, amt, note in JUL_VISITS:
        put(wb["Jul-2026"], find(wb["Jul-2026"], "visit", d, f, n), amt, note)
    for sec, d, f, n, amt, note in AUG:
        put(wb["Aug-2026"], find(wb["Aug-2026"], sec, d, f, n), amt, note)
    wb.save(path)
    print("amounts filled:", len(JUL_VISITS), "July visits,", len(AUG), "August rows")


if __name__ == "__main__":
    main(sys.argv[1])
