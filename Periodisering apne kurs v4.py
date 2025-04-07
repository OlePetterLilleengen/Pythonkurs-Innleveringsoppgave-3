import pandas as pd
from tkinter import Tk, simpledialog, filedialog

def generate_utput_rader(rad, bilagsnr, aar, maanednr):
    jan = str(aar) + "0131"
    feb = str(aar) + "0228"
    mars = str(aar) + "0331"
    april = str(aar) + "0430"
    mai = str(aar) + "0531"
    juni = str(aar) + "0630"
    aug = str(aar) + "0831"
    sept = str(aar) + "0930"
    okt = str(aar) + "1031"
    nov = str(aar) + "1130"
    des = str(aar) + "1231"
    faktbelop = -rad['belopsum']
    if maanednr == 1:
        perbelop = faktbelop / -6
        dates = [jan, jan, feb, mars, april, mai, juni]  * 2
        kontoer = [3295] * 7 + [2965] * 7
        belop = [faktbelop] + [perbelop] * 6 + [-faktbelop] + [-perbelop] * 6
    elif maanednr == 2:
        perbelop = faktbelop / -5
        dates = [feb, feb, mars, april, mai, juni] * 2
        kontoer = [3295] * 6 + [2965] * 6
        belop = [faktbelop] + [perbelop] * 5 + [-faktbelop] + [-perbelop] * 5
    elif maanednr == 8:
        perbelop = faktbelop / -5
        dates = [aug, aug, sept, okt, nov, des] * 2
        kontoer = [3295] * 6 + [2965] * 6    
        belop = [faktbelop] + [perbelop] * 5 + [-faktbelop] + [-perbelop] * 5
    elif maanednr == 9:
        perbelop = faktbelop / -4
        dates = [sept, sept, okt, nov, des] * 2
        kontoer = [3295] * 5 + [2965] * 5 
        belop = [faktbelop] + [perbelop] * 4 + [-faktbelop] + [-perbelop] * 4

    

    utputs = []
    for i in range(len(dates)):
        description = f"Periodisering pr. nr. {rad['Prosjektnummer']} {str(int(rad['avdelingsum']))}"
        utput = {
            'utfilformat': 'GBAT10',
            'utbilagsnr': bilagsnr,
            'utbilagsdato': dates[i],
            'utbilagstype': 1,
            'utblank1': 1,
            'utbilagsaar': aar,
            'utkonto': kontoer[i],
            'utmomskode': '0',
            'utbelop': belop[i],
            'utkundenr': 0,
            'utleverandornr': 0,
            'utkundevnavn': '',
            'utkundelevadresse': '',
            'utkundelevpostnr': '',
            'utkundelevby': '',
            'utblank2': '',
            'utblank3': '',
            'utblank4': '',
            'utblank5': '',
            'utblank6': '',
            'utbeskrivelse': description,
            'utputblank7': '',
            'utputblank8': '',
            'utprosjektnr': rad['Prosjektnummer'],
            'utavdeling': str(int(rad['avdelingsum'])),
            'utblank9': 1,
            'utblank10': 'T',
            'utbruttobelop': belop[i]
        }
        utputs.append(utput)
    return utputs

def main():
    root = Tk()
    root.withdraw()  # Skjuler hovedvinduet for tkinter

    # Samle brukerinput via dialoger
    aar = simpledialog.askinteger("Input", "Skriv inn år:", parent=root)
    maanednr = simpledialog.askinteger("Input", "Skriv inn månedsnr:", parent=root)
    bilagsnr = simpledialog.askinteger("Input", "Skriv inn bilagsnummer:", parent=root)

    # default_directory = "C:/Users/OlePetterLilleengen/K2Kompetanse/K2 Kompetanse sharepoint - Dokumenter/K2 Kompetanse/Okonomi/Regnskap/Periodisering/"

    # Velge input-fil
    file_path = filedialog.askopenfilename(
        title="Velg en Excel-fil",
        filetypes=[("Excel files", "*.xlsx")],
   #     initialdir=default_directory
    )
    if not file_path:
        print("Ingen fil ble valgt.")
        return

    # Lese inn data
    input_df = pd.read_excel(file_path)
    filtered_df = input_df[~input_df['Avdelingsnummer'].isin([252, 102, 900])]
    sumprosjekt = filtered_df.groupby('Prosjektnummer').agg({
        'Prosjektnavn': 'first',
        'Avdelingsnummer': 'first',
        'Beløp': 'sum'
    }).rename(columns={'Prosjektnavn': 'prosjektnavnsum', 'Avdelingsnummer': 'avdelingsum', 'Beløp': 'belopsum'}).reset_index()
    sumprosjekt = sumprosjekt[sumprosjekt['belopsum'] != 0]
    
    print(sumprosjekt)

    utputs = []
    # Loop gjennom hver rad i sumprosjekt og generer utput rader basert på månedsnummer
    for _, rad in sumprosjekt.iterrows():
        utputs.extend(generate_utput_rader(rad, bilagsnr, aar, maanednr))

    utput_df = pd.DataFrame(utputs)

# Konverter beløp til strenger med komma som desimalskille
    utput_df['utbelop'] = utput_df['utbelop'].apply(lambda x: f"{x:.5f}".replace('.', ','))
    utput_df['utbruttobelop'] = utput_df['utbruttobelop'].apply(lambda x: f"{x:.5f}".replace('.', ','))

# La brukeren velge hvor og under hvilket navn filen skal lagres
    
    file_path = filedialog.asksaveasfilename(
        title="Lagre som CSV",
        filetypes=[("CSV files", "*.csv")],
        defaultextension=".csv",
#20        initialdir=default_directory
    )
    if file_path:
        utput_df.to_csv(file_path, index=False, sep=';', header=False)
        print("Filen er lagret til:", file_path)
    else:
        print("Ingen fil ble lagret.")
    
        

if __name__ == "__main__":
    main()
