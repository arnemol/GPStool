import pandas as pd
import os
from tkinter import Tk, filedialog

# Functie om het txt-bestand te selecteren
def selecteer_bestand():
    """
    Opent een bestandsdialoog om een GPS-gegevensbestand te selecteren.
    """
    root = Tk()
    root.withdraw()  # Verberg het hoofdvenster
    bestandspad = filedialog.askopenfilename(title="Selecteer GPS data bestand", 
                                             filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
    root.destroy()
    return bestandspad

# Functie om het txt-bestand in te lezen en te verwerken
def lees_en_verwerk_data(bestandspad):
    """
    Leest en verwerkt de GPS-gegevens uit het geselecteerde tekstbestand.
    """
    try:
        # Kolomnamen zoals ze in het bestand worden verwacht
        kolommen = ['FID', 'VID', 'X', 'Y', 'Z', 'Nauwkeurigheid', 'putnummer']
        
        # Data inlezen met specifieke delimiter en encoding
        df = pd.read_csv(bestandspad, delimiter=';', names=kolommen, skiprows=1, encoding='utf-16')

        # Voeg een lege kolom 'Melding' toe voor toekomstige foutmeldingen
        df['Melding'] = ''

        # Controleer of de 'Nauwkeurigheid' kolom correct is ingelezen
        if 'Nauwkeurigheid' not in df.columns:
            raise ValueError("De kolom 'Nauwkeurigheid' is niet gevonden in het ingelezen bestand.")
        
        # Tellen van rijen met 'Nauwkeurigheid' > 0,02
        count_nauwkeurigheid_filtered = df[df['Nauwkeurigheid'] > 0.02].shape[0]

        # Voeg meldingen toe aan rijen met 'Nauwkeurigheid' > 0,02
        df.loc[df['Nauwkeurigheid'] > 0.02, 'Melding'] = (
            f'Waarde hoger dan 0,02 gedetecteerd, waarde gefilterd. Totaal: {count_nauwkeurigheid_filtered} waarden'
        )

        # Filter de data om alleen rijen met 'Nauwkeurigheid' <= 0,02 te behouden
        filtered_df = df[df['Nauwkeurigheid'] <= 0.02]

        # Bereken gemiddelden per putnummer
        gemiddelden = filtered_df.groupby('putnummer').agg({
            'X': 'mean',
            'Y': 'mean',
            'Z': 'mean'
        }).reset_index()

        # Voeg een kolom 'Nauwkeurigheid' toe aan de gemiddelden DataFrame
        gemiddelden['Nauwkeurigheid'] = 0.02  # Aangezien alleen metingen met <= 0.02 worden meegenomen

        # Voeg een kolom toe voor afwijkingen en meldingen
        afwijkingen = []

        for putnummer, groep in df.groupby('putnummer'):
            gemiddeld = gemiddelden[gemiddelden['putnummer'] == putnummer].iloc[0]
            afwijkingen_groep = groep.copy()
            
            afwijkingen_groep['X_afwijking'] = abs(groep['X'] - gemiddeld['X'])
            afwijkingen_groep['Y_afwijking'] = abs(groep['Y'] - gemiddeld['Y'])
            afwijkingen_groep['Z_afwijking'] = abs(groep['Z'] - gemiddeld['Z'])
            
            # Tellen van rijen met afwijkingen groter dan opgegeven eenheden
            count_afwijkingen = afwijkingen_groep[
                (afwijkingen_groep['X_afwijking'] > 10) | 
                (afwijkingen_groep['Y_afwijking'] > 10) | 
                (afwijkingen_groep['Z_afwijking'] > 3)
            ].shape[0]

            # Markeer rijen met afwijkingen groter dan opgegeven eenheden
            afwijkingen_groep['Melding'] = afwijkingen_groep.apply(
                lambda row: f'Waarde afgeweken meer dan toegestane eenheden, niet meegenomen in gemiddelde. Totaal: {count_afwijkingen} waarden'
                if (row['X_afwijking'] > 10) or (row['Y_afwijking'] > 10) or (row['Z_afwijking'] > 3) 
                else row['Melding'], axis=1
            )
            
            afwijkingen.append(afwijkingen_groep)

        afwijkingen_df = pd.concat(afwijkingen)

        # Voeg meldingen toe aan de gefilterde DataFrame voor afwijkingen
        filtered_final_df = afwijkingen_df[~(
            (afwijkingen_df['X_afwijking'] > 10) |
            (afwijkingen_df['Y_afwijking'] > 10) |
            (afwijkingen_df['Z_afwijking'] > 3)
        )]

        # Voeg de meldingen toe aan de resultaten
        meldingen = pd.concat([
            df.groupby('putnummer')['Melding'].apply(lambda x: ', '.join(x.unique())).reset_index(),
            afwijkingen_df.groupby('putnummer')['Melding'].apply(lambda x: ', '.join(x.unique())).reset_index()
        ]).groupby('putnummer').agg({'Melding': lambda x: ', '.join(x.unique())}).reset_index()

        final_results = pd.merge(gemiddelden, meldingen, on='putnummer', how='left')

        # Voeg een kolom toe voor sortering
        final_results['Is_mv'] = final_results['putnummer'].apply(lambda x: 1 if x.endswith('-mv') else 0)

        # Sorteer op basis van 'Is_mv' zodat '-mv' records onderaan komen
        final_results = final_results.sort_values(by=['Is_mv', 'putnummer'])

        # Verwijder de tijdelijke kolom 'Is_mv'
        final_results = final_results.drop(columns=['Is_mv'])

        # Round X, Y, Z coordinates to 3 decimal places
        final_results[['X', 'Y', 'Z']] = final_results[['X', 'Y', 'Z']].round(3)

        # Kolom volgorde bepalen voor output
        kolom_volgorde = [
            'putnummer', 'X', 'Y', 'Z', 
            'Nauwkeurigheid', 'Melding'
        ]
        final_results = final_results[kolom_volgorde]

        # Debug: Print de resultaten DataFrame voor schrijven naar bestand
        print("Final Results DataFrame voor schrijven naar bestand:")
        print(final_results.head())

        return final_results

    except Exception as e:
        print(f"Er trad een fout op bij het verwerken van de data: {e}")
        return None

# Functie om resultaten naar een Excel-bestand te schrijven en het bestand te openen
def schrijf_resultaten_naar_bestand(resultaten, uitvoerpad):
    """
    Schrijft de verwerkte resultaten naar een Excel-bestand en opent dit bestand.
    """
    try:
        if resultaten is not None and not resultaten.empty:
            # Schrijf de resultaten naar het Excel-bestand
            resultaten.to_excel(uitvoerpad, index=False, engine='openpyxl')
            
            # Debug: Print pad van het uitvoerbestand
            print(f"Resultaten zijn geschreven naar: {uitvoerpad}")

            os.startfile(uitvoerpad)  # Open het bestand automatisch (alleen op Windows)
        else:
            print("Geen resultaten om naar bestand te schrijven.")

    except Exception as e:
        print(f"Er trad een fout op bij het schrijven van de resultaten naar het bestand: {e}")

# Hoofdprogramma
if __name__ == "__main__":
    bestandspad = selecteer_bestand()

    if bestandspad:  # Controleer of er een bestand is geselecteerd
        resultaten = lees_en_verwerk_data(bestandspad)

        if resultaten is not None:
            uitvoerpad = os.path.splitext(bestandspad)[0] + '_resultaten.xlsx'
            schrijf_resultaten_naar_bestand(resultaten, uitvoerpad)
        else:
            print("Geen resultaten om weer te geven.")
    else:
        print("Geen bestand geselecteerd.")
