"""Add the rows from the Nov-2026 uploads that were NOT already in the master workbook.

Run on the design-system base (before scripts/add_radiology_sections.py):
    python scripts/merge_uploads_nov2026.py <master.xlsx>

Every row below was checked against the master by date + file number and by name;
everything else in the uploads was already present (see PR description).
"""
import re
import sys
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.comments import Comment

sys.path.insert(0, __file__.rsplit("/", 1)[0])
import design_system as ds  # noqa: E402

# sheet -> rows (date, file, name, consultation[, note]). July: all rows of the TrueLife file
# (FAIZ_PTS_DATA_TRUELIFE_NEW.xlsx, sheet "Jul 26"); the 2 Tawoor Khan visits were already in the
# master (from faiz_marketing_AUG_data). Names: master spelling for known file numbers, run-together
# names split, first + second name only when unclear (repo naming rule). Consultation = doctors as given.
NEW_VISITS = {
    "Jul-2026": [
        ('01.07.2026', 'TVIP009213588', 'Tawoor Khan', 'Physio'),
        ('01.07.2026', 'TVIP00980784', 'Sumithra Venkatesan', 'Dr Nandini'),
        ('01.07.2026', 'TVIP01013249', 'Ehsan Ullah', 'Dr Feras', "Name in source: 'Ehsanu Llaherfanu Llah'."),
        ('02.07.2026', 'TVIP01013128', 'Muhammed Ahyaan', 'Dr Bassima'),
        ('04.07.2026', 'TVIP009213588', 'Tawoor Khan', 'Physio'),
        ('04.07.2026', 'TVIP01013800', 'Mansoor Ali Korankulangara', 'Dr Mozammil', "Name in source: 'Mansoorali Korankulangara'."),
        ('05.07.2026', 'TVIP01013128', 'Muhammed Ahyaan', 'Dr Bassima'),
        ('06.07.2026', 'TVIP00829926', 'Leela Rairoth', 'Dr Afsheen, Dr Ihab'),
        ('13.07.2026', 'TVIP01015079', 'Abdul Gafoor Kummalil', 'Dr Roaa'),
        ('14.07.2026', 'TVIP00830036', 'Bushra Bilal Muhammad', 'Dr Nour'),
        ('14.07.2026', 'TVIP01015139', 'Eassa Karuvante', 'Dr Agnel'),
        ('16.07.2026', 'TVIP00902268', 'Muhammed Ali Ponneth', 'Dr Afsheen'),
        ('18.07.2026', 'TVIP00658599', 'Maheen Abdulrahuman', 'Dr Imad'),
        ('18.07.2026', 'TVIP00902268', 'Muhammed Ali Ponneth', 'Dr Feras'),
        ('19.07.2026', 'TVIP00654762', 'Shahzad Afsar', 'Dr Afsheen, Dr Nandini'),
        ('19.07.2026', 'TVIP00961592', 'Muhammad Zaroob', 'Dr Afsheen'),
        ('20.07.2026', 'TVIP01015196', 'Adel Fares Afifi Ibrahim Afifi', 'Dr Higazy; Dr Mofeed, Dr Ihab, Dr Alamgir', "Name in source: 'Adel Fares Afifiibrahimafifi'."),
        ('20.07.2026', 'TVIP01015969', 'Rubel Miah Suruj Miah', 'Dr Afsheen'),
        ('20.07.2026', 'TVIP01015970', 'Mohammed Sabbir Ahmed', 'Dr Afsheen'),
        ('21.07.2026', 'TVIP00658599', 'Maheen Abdulrahuman', 'Dr Imad, Dr Nandini'),
        ('21.07.2026', 'TVIP00979556', 'Zayan Hamad Maheen', 'Dr Roaa, Dr Nandini'),
        ('21.07.2026', 'TVIP01016042', 'Shihabudheen Palakkal', 'Dr Nighet'),
        ('25.07.2026', 'TVIP00902268', 'Muhammed Ali Ponneth', 'Dr Feras, Dr Ali'),
        ('25.07.2026', 'TVIP00973129', 'Hiba Kunnummal', 'Dr Walaa'),
        ('25.07.2026', 'TVIP01013128', 'Muhammed Ahyaan', 'Dr Bassima'),
        ('25.07.2026', 'TVIP01015969', 'Rubel Miah Suruj Miah', 'Dr Roaa'),
        ('25.07.2026', 'TVIP01015970', 'Mohammed Sabbir Ahmed', 'Dr Roaa'),
        ('30.07.2026', 'TVIP00615797', 'Abdurahman Olimeah', 'Dr Dimeh'),
        ('30.07.2026', 'TVIP00865589', 'Nijam Uddin Mubarok Ali', 'Dr Mozammil, Dr Sibtain, Dr Radi', "Name in source: 'Nijam Uddinmubarok Ali'."),
        ('30.07.2026', 'TVIP00925005', 'Sanitha Valiyakath', 'Dr Shiraz, Dr Nighet', "Name in source: 'Sanithavaliyakath Sulaiman Valiyaka'."),
        ('30.07.2026', 'TVIP00948643', 'Muhammed Hameed', 'Dr Dimeh'),
        ('30.07.2026', 'TVIP00963637', 'Nahala Variammadathil', 'Dr Dimeh', "Name in source: 'Nahalavariammadathil'."),
        ('30.07.2026', 'TVIP00974777', 'Saber Rezk Elhussein', 'Dr Roaa', "Name in source: 'Saberrezk Elhussein'."),
        ('30.07.2026', 'TVIP01001897', 'Mohammed Ameer Puzhangaraillath', 'Dr Nighet', "Name in source: 'Mohammed Ameerpuzhangaraillath'."),
        ('30.07.2026', 'TVIP01017072', 'Chandran Velath', 'Dr Mustafa, Dr Radi, Dr Abdulhaleem', "Name in source: 'Chandranvelath Vallikkateroma Menon'."),
        ('30.07.2026', 'TVIP01017073', 'Ramakrishnan Melekadungoth', 'Dr Mustafa, Dr Higazy, Dr Shiraz', "Name in source: 'Ramakrishnanmelekadungoth'."),
        ('30.07.2026', 'TVIP01017075', 'Hammad Hussain Janjua', 'Dr Mustafa, Dr Abdulhaleem', "Name in source: 'Hammad Hussainjanjua'."),
        ('30.07.2026', 'TVIP01017076', 'Jamghat Gul', 'Dr Mustafa, Dr Alaa', "Name in source: 'Jamghatgularmani'."),
        ('30.07.2026', 'TVIP01017118', 'Rodello Olazo Santos', 'Dr Nighet', "Name in source: 'Rodello Olazosantos'."),
        ('30.07.2026', 'TVIP01017120', 'Brudi Solamon Korattikunnan', 'Dr Nighet', "Name in source: 'Brudisolamon Korattikunnan'."),
        ('30.07.2026', 'TVIP01017122', 'Ibrahim Thayyil Hamza', 'Dr Mozammil', "Name in source: 'Ibrahim Thayyilhamza'."),
        ('30.07.2026', 'TVIP01017126', 'Jeganathan Perumal', 'Dr Mozammil', "Name in source: 'Jeganathanperumalperumal'."),
        ('30.07.2026', 'TVIP01017132', 'Haneefa Thekkumbath', 'Dr Mozammil', "Name in source: 'Haneefathekkumbath'."),
        ('30.07.2026', 'TVIP01017134', 'Mohammad Asim Nasimur Rehaman', 'Dr Mozammil', "Name in source: 'Mohammad Asimnasimur Rehaman'."),
        ('30.07.2026', 'TVIP01017151', 'Balu Chandrashekar', 'Dr Roaa, Dr Abdulhaleem, Dr Ali, Dr Jimshad', "Name in source: 'Baluchandirasekarchandirasekar'."),
        ('30.07.2026', 'TVIP01017168', 'Jishnu Thamarassery', 'Dr Roaa, Dr Ali', "Name in source: 'Jishnuthamarassery'."),
        ('30.07.2026', 'TVIP01017170', 'Renjith Raju', 'Dr Roaa, Dr Tagreed', "Name in source: 'Renjithraju Raju Pazhayidath'."),
        ('30.07.2026', 'TVIP01017171', 'Unnikrishnan Thelasseri Parambil', 'Dr Roaa, Dr Tagreed', "Name in source: 'Unnikrishnanthelasseri Parambil'."),
        ('30.07.2026', 'TVIP01017173', 'Vijayakumar Arumugam', 'Dr Roaa, Dr Agnel, Dr Ali', "Name in source: 'Vijayakumararumugamarumugam'."),
        ('30.07.2026', 'TVIP01017181', 'Naze Ahmed Rashed', 'Dr Mozammil', "Name in source: 'Nazeahmed Rashed Ahmed'."),
        ('30.07.2026', 'TVIP01017184', 'Junaid Ahmed Jamaluddin', 'Dr Mozammil, Dr Agnel', "Name in source: 'Junaidahmed Jamaluddin'."),
        ('30.07.2026', 'TVIP01017186', 'Sahid Mahammad Nur Mahammad', 'Dr Mozammil', "Name in source: 'Sahidmahammadnur Mahammad'."),
        ('30.07.2026', 'TVIP01017188', 'Abdul Qayum Zar Ur Rehman', 'Dr Roaa, Dr Zoubi', "Name in source: 'Abdul Qayumzar UR Rehman'."),
        ('30.07.2026', 'TVIP01017198', 'Anas Makkikaran', 'Dr Dimeh', "Name in source: 'Anas Makkikaranabdul'."),
        ('30.07.2026', 'TVIP01017200', 'Nuha Bint Mohammed Ameer', 'Dr Sameh, Dr Agnel', "Name in source: 'Nuhabint Mohammed Ameermohammed Ameer'."),
        ('30.07.2026', 'TVIP01017202', 'Huda Bint Mohammed Ameer', 'Dr Sameh', "Name in source: 'Huda Bintmohammed Ameer'."),
        ('30.07.2026', 'TVIP01017204', 'Wafa Bint Mohammed Ameer', 'Dr Sameh', "Name in source: 'Wafabintmohammed Ameer'."),
        ('30.07.2026', 'TVIP01017206', 'Rana Muhammad Zubair', 'Dr Dimeh, Dr Agnel', "Name in source: 'Rana Muhammadzubair'."),
        ('30.07.2026', 'TVIP01017231', 'Mohamed Fazil Erambath Alassan', 'Dr Mozammil', "Name in source: 'Mohamedfazil Erambath Alassan'."),
        ('30.07.2026', 'TVIP01017232', 'Riyas Pilakkot Mohamed Kutty', 'Dr Mozammil, Dr Alaa', "Name in source: 'Riyas Pilakkotmohamed Kutty'."),
        ('30.07.2026', 'TVIP01017234', 'Mohammed Shahid Ullah Mohammed', 'Dr Dimeh, Dr Agnel', "Name in source: 'Mohammedshahid Ullah Mohammed'."),
        ('30.07.2026', 'TVIP01017250', 'Mohammedkutty Edathadathil', 'Dr Roaa'),
        ('30.07.2026', 'TVIP01017253', 'Abdul Rahim', 'Dr Roaa, Dr Sibtain', "Name in source: 'Abdul Rahimmaviladaththaila'."),
        ('30.07.2026', 'TVIP01017260', 'Soukath Ali Kuthubuddin', 'Dr Roaa, Dr Bassima', "Name in source: 'Soukathalikuthubuddin'."),
        ('30.07.2026', 'TVIP01017276', 'Awatef Ramadan Hussein', 'Dr Roaa', "Name in source: 'Awateframadan Hussein'."),
        ('30.07.2026', 'TVIP01017279', 'Mohamed Saber Rezk Elhusseini', 'Dr Roaa', "Name in source: 'Mohamedsaber Rezk Elhusseini'."),
        ('30.07.2026', 'TVIP01017284', 'Thomas Mathew', 'Dr Roaa, Dr Agnel', "Name in source: 'Thomasmathew'."),
        ('30.07.2026', 'TVIP01017285', 'Ajosh David Suresh', 'Dr Roaa, Dr Agnel', "Name in source: 'Ajosh Davidsuresh'."),
        ('30.07.2026', 'TVIP01017286', 'Swaraj Kurup Chandrasekharan', 'Dr Dimeh, Dr Ali', "Name in source: 'Swarajkurup Chandrasekharan'."),
        ('30.07.2026', 'TVIP01017287', 'Sultan Ahmed Nizam Uddin', 'Dr Roaa', "Name in source: 'Sultan Ahmednizam Uddin'."),
        ('30.07.2026', 'TVIP01017289', 'Firos Cheriya', 'Dr Dimeh'),
        ('30.07.2026', 'TVIP01017290', 'Eiwa Yaliq', 'Dr Dimeh', "Name in source: 'Eiwa Yaliqfiros Firos Cheriy'."),
        ('30.07.2026', 'TVIP01017291', 'Hood Yaliq Firos', 'Dr Dimeh', "Name in source: 'Hood Yaliqfiros'."),
        ('30.07.2026', 'TVIP01017292', 'Mohammed Harunur Rashid Delwar', 'Dr Dimeh', "Name in source: 'Mohammedharunur Rashid Delwar'."),
        ('30.07.2026', 'TVIP01017293', 'Ameer Arinjikkal', 'Dr Dimeh', "Name in source: 'Ameerarinjikkal'."),
        ('30.07.2026', 'TVIP01017288', 'Abdul Rahiman Badiadka', 'Dr Dimeh', "File no. written 'vTVIP01017288' in source. Name in source: 'Abdul Rahimanbadiadka'."),
        ('31.07.2026', 'TVIP00604064', 'Sabil Ullah Hajat Khan', 'Dr Roaa, Dr Ali'),
        ('31.07.2026', 'TVIP01017118', 'Rodello Olazo Santos', 'Dr Agnel; Dr Nighet', "Name in source: 'Rodello Olazosantos'."),
        ('31.07.2026', 'TVIP01017311', 'Pramod Radhakrishna Pillai', 'Dr Dimeh', "Source gives TVIP01017288 (same as Abdul Rahiman Badiadka); the master's Aug record for this patient is TVIP01017311. Name in source: 'Pramodradhakrishnapillai'."),
        ('31.07.2026', 'TVIP01017297', 'Vinoth Kumar Murugan', 'Dr Afsheen'),
        ('31.07.2026', 'TVIP01017307', 'Rusna Muchamad Rambang', 'Dr Roaa, Dr Agnel'),
        ('31.07.2026', 'TVIP01017312', 'Sureshkumar Chandrasekhara Pillai', 'Dr Nighet', "Name in source: 'Sureshkumar Chandrasekharapillai'."),
        ('31.07.2026', 'TVIP01017324', 'Mohammed Hadi', 'Dr Roaa', "Name in source: 'Mohammed Hadilellath Parambil'."),
        ('31.07.2026', 'TVIP01017325', 'Mohammed Shafeeq Thazhathodiyil', 'Dr Roaa'),
        ('31.07.2026', 'TVIP01017326', 'Muhammed Shafi Thazhathodi', 'Dr Roaa'),
        ('31.07.2026', 'TVIP01017331', 'Hamid Mohammed Qasem Ahmed', 'Dr Mustafa', "Name in source: 'Hamid Mohammed Qasemahmed'."),
        ('31.07.2026', 'TVIP01017333', 'Abdul Basith Manalodi', 'Dr Mustafa'),
        ('31.07.2026', 'TVIP01017334', 'Marwa Ameen', 'Dr Mustafa'),
        ('31.07.2026', 'TVIP01017335', 'Turki Qasem', 'Dr Mustafa; Dr Zafer'),
    ],
    # (Nihala Sherin, 6 Aug, TVIP00898744 is a radiology referral per the August-2026 file:
    #  she is in the Aug-2026 radiology section, not here.)
}
# Source: Spectrum_Referrals_Town_Centre_Branch.xlsx - the only 'Done' referral not already
# in the Oct-2026 radiology list (others matched by name at their scan date).
OCT_RADIOLOGY = ("24.10.2026", "18639", "Mohamed Abbas", "CT LEFT FOOT",
                 "From Spectrum Referrals - Town Centre Branch (status: Done). File no. is the "
                 "Town Centre EMR #; bill amount not given in source.")

REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def shift_formula(f, idx, only_sheet=None):
    """Shift row refs >= idx down by one (refs to `only_sheet` only, when given)."""
    if only_sheet is None:
        return REF.sub(lambda m: m.group(1) + m.group(2) + m.group(3)
                       + str(int(m.group(4)) + (int(m.group(4)) >= idx)), f)
    pat = re.compile(r"('%s'!)(\$?[A-Z]{1,3}\$?)(\d+)" % re.escape(only_sheet))
    return pat.sub(lambda m: m.group(1) + m.group(2) + str(int(m.group(3)) + (int(m.group(3)) >= idx)), f)


def insert_row(wb, ws, idx, style_row):
    merges = [r.coord for r in ws.merged_cells.ranges]
    heights = {k: d.height for k, d in ws.row_dimensions.items() if d.height is not None}
    for m in merges:
        ws.unmerge_cells(m)
    ws.insert_rows(idx)
    for m in merges:
        ws.merge_cells(shift_formula(m, idx))
    for k in list(ws.row_dimensions.keys()):
        ws.row_dimensions[k].height = None
    for k, h in heights.items():
        ws.row_dimensions[k + (k >= idx)].height = h
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = shift_formula(c.value, idx)
    for other in wb.worksheets:
        if other is ws:
            continue
        for row in other.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and f"'{ws.title}'!" in c.value:
                    c.value = shift_formula(c.value, idx, ws.title)
    src = style_row + (style_row >= idx)
    for col in range(1, ws.max_column + 1):
        s, t = ws.cell(src, col), ws.cell(idx, col)
        t.font, t.fill, t.border = copy(s.font), copy(s.fill), copy(s.border)
        t.alignment, t.number_format = copy(s.alignment), s.number_format
    ws.row_dimensions[idx].height = heights.get(style_row, ds.ROW_H["data"])


def renumber(ws, first, last):
    for i, r in enumerate(range(first, last + 1), 1):
        ws.cell(r, 1).value = i


def add_visits(wb, sheet, rows):
    ws = wb[sheet]
    total = ws.max_row
    spacer = total - 1
    for d, fno, name, svc, *note in rows:
        dd = datetime.strptime(d, "%d.%m.%Y")
        data = list(range(4, spacer))
        used = [r for r in data if ws.cell(r, 2).value is not None]
        later = [r for r in used if ws.cell(r, 2).value > dd]
        nxt = (used[-1] + 1) if used else 4
        if not later and nxt < spacer and ws.cell(nxt, 2).value is None:
            r = nxt  # blank template row right after the last entry
        else:
            r = later[0] if later else nxt
            insert_row(wb, ws, r, 4)
            total += 1
            spacer += 1
            data = list(range(4, spacer))
        for col, v in zip("BCDE", (dd, fno, name, svc)):
            ws[f"{col}{r}"] = v
        for col in "ABCDE":
            ds.style_data(ws[f"{col}{r}"], col)
        if note:
            ws[f"D{r}"].comment = Comment(note[0], "Claude")
        used = [x for x in data if ws.cell(x, 2).value is not None]
    renumber(ws, 4, used[-1])
    for r in data:
        if ws.cell(r, 2).value is None:
            ws.cell(r, 1).value = None
    # rows appended after the template end are outside the old COUNTA range: rewrite it
    total = next(r for r in range(4, ws.max_row + 1)
                 if str(ws.cell(r, 1).value or "").startswith("TOTAL PATIENTS"))
    ws.cell(total, 5).value = f"=COUNTA(B4:B{total - 2})"


def add_oct_radiology(wb):
    ws = wb["Oct-2026"]
    d, fno, name, exam, note = OCT_RADIOLOGY
    dd = datetime.strptime(d, "%d.%m.%Y")
    hdr = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Sl No" and r > 3)
    rows = [r for r in range(hdr + 1, ws.max_row + 1) if isinstance(ws.cell(r, 2).value, datetime)]
    idx = next(r for r in rows if ws.cell(r, 2).value > dd)
    insert_row(wb, ws, idx, rows[0])
    for col, v in zip("BCDE", (dd, fno, name, exam)):
        ws[f"{col}{idx}"] = v
    ws[f"D{idx}"].comment = Comment(note, "Claude")
    renumber(ws, hdr + 1, rows[-1] + 1)


def main(path):
    wb = openpyxl.load_workbook(path)
    for sheet, rows in NEW_VISITS.items():
        add_visits(wb, sheet, rows)
    add_oct_radiology(wb)
    s = wb["Summary"]
    s["A19"] = s["A19"].value.replace("Jul 2026 has no data yet (blank template)",
                                      "Jul 2026 from the TrueLife and marketing AUG files")
    wb.save(path)


if __name__ == "__main__":
    main(sys.argv[1])
