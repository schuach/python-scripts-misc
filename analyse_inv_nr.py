# Script, das aus einer Liste von Inventarnummern eine Liste der fehlenden ausgibt
import pyperclip
import easygui

# Die Liste der zulässigen Präfixe und das padding
prefixes = {
    "2018-CJS-":	3,
    "2018-CP-":	3,
    "2018-ANM-":	3,
    "2018-GK-":	3,
    "2018-GR-":	4,
    "2018-GSL-":	4,
    "2018-GSLI-":	3,
    "2018-GU-":	3,
    "2018-GVA-":	4,
    "2018-HL-":	2,
    "2018-GVF-":	2,
    "2018-GVS-":	2,
    "2018-HDW-":	2,
    "2018-INIG-":	2,
    "2018-IPS-":	2,
    "2018-ISIS-":	2,
    "2018-K-":	3,
    "2018-L-":	3,
    "2018-M-":	3,
    "2018-MA-":	3,
    "2018-MAO-":	2,
    "2018-MI-":	3,
    "2018-ZK-":	2,
    "2018-MSO-":	3,
    "2018-MR-":	3,
    "2018-MTHK-":	3,
    "2018-RBO-":	2,
    "2018-RC-":	2,
    "2018-RZ-":	3,
    "2018-SB-":	2,
    "2018-SBO-":	2,
    "2018-SC-":	2,
    "2018-SD-":	2,
    "2018-SE-":	0,
    "2018-SF-":	2,
    "2018-SFE-":	2,
    "2018-SFE-F-":	2,
    "2018-SGA-":	0,
    "2018-SGB-":	0,
    "2018-SGC-":	0,
    "2018-SGD-":	0,
    "2018-SGF-":	2,
    "2018-SGFD-":	3,
    "2018-SGG-":	0,
    "2018-SGH-":	0,
    "2018-SGP-":	0,
    "2018-SGPT-":	0,
    "2018-SGT-":	0,
    "2018-SH-":	4,
    "2018-SPH-":	2,
    "2018-STO-":	2,
    "2018-TRSP-":	3,
    "2018-W-":	2,
    "2018-WEGC-":	3,
    "2018-ZANT-":	3,
    "2018-HB-":	5,
    "2018-GA-":	3,
    "2018-GD-":	3,
    "2018-GG-":	4,
    "2018-IP-":	3,
    "2018-O-":	3,
    "2018-RBA-":	2,
    "2018-RSW-":	4,
    "2018-THEO-":	4,
    "2018-TRSPKI-":	0,
    "2018-TRSPPL-":	0,
    "2018-LZ-":	2,
}

# die liste aus der Zwischenablage holen
def get_instr():
    """Gets contents from clipboard and checks them for validity. Returns a
    string.
    """
    msg = """Bitte die Spalte in die Zwischenablage Kopieren.

Klicken Sie dazu in Excel auf den Spaltenkopf und drücken Sie Strg-C."""

    while True:
        reply = easygui.buttonbox(msg, choices=["OK", "Abbrechen"])

        if reply == "Abbrechen":
            quit()
        else:
            instr = pyperclip.paste()

        # check if data is valid -- if not, ask again
        if "Inventarnummer" not in instr:
            msg = """Etwas scheint nicht zu stimmen. Bitte versuchen Sie es nocheinmal.

Klicken Sie in Excel auf den Spaltenkopf und drücken Sie Strg-C."""
            continue
        else:
            return instr


inlist = get_instr().split("\n")

def get_prefix():
    prefix = ''
    msg = 'Bitte das Präfix auswählen:'
    choices = sorted(prefixes.keys())
    prefix = easygui.choicebox(msg, "Inventarnummernkreis", choices)
    return prefix


# die vorhandenen nummern als int in eine Liste schreiben
prefix = get_prefix()
padding = prefixes[prefix]
inv_int = []
errors = []
len_pref = len(prefix)
for nr in inlist:
    if "Inventarnummer" in nr:
        continue
    elif nr == "":
        continue
    elif not nr.startswith(prefix):
        errors.append(nr)
    else:
        try:
            inv_int.append(int(nr[len_pref:]))
        except:
            errors.append(nr)

highest_number = max(inv_int)

# Die fehlenden Nummern finden
missing_nrs = []
for i in range(1, highest_number):
    if i not in inv_int:
        missing_nrs.append(i)

# Duplikate finden
seen = {}
dupes = []
for i in inlist:
    if i not in seen:
        seen[i] = 1
    else:
        if seen[i] == 1:
            dupes.append(i)
        seen[i] += 1


# Die Ausgabedatei schreiben
filename = f"Analyse_Inventarnummern_{prefix[:-1]}.txt"
with open(filename, "w") as outfile:
    outfile.write(f"Inventarnummern-Check für {prefix}: \n")
    outfile.write(f"Inventarnummern vorhanden: {len(errors) + len(inv_int)}\n")
    outfile.write(f"\nDavon strukturell falsch: {len(errors)}\n")
    for error in errors:
        outfile.write(error + "\n")

    outfile.write(f"\n\n{len(missing_nrs)} fehlende Inventarnummern:\n")
    for nr in missing_nrs:
        outfile.write(f"{prefix}{nr:0{padding}}\n")

    outfile.write(f"\n\n{len(dupes)} mehrfach vergebene Inventarnummern:\n")
    for dupe in dupes:
        outfile.write(f"{dupe}, {seen[dupe]}x\n")
