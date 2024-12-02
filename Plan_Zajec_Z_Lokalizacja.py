import pandas as pd
from datetime import datetime, timedelta
import os
from ics import Calendar, Event

file_path = r'<<SCIEZKA DO PLIKU .EXCEL>>'
df = pd.read_excel(file_path, sheet_name='<<> NAZWA ARKUSZA >>')

df_filtered = df.iloc[5:].reset_index(drop=True)
df_filtered.columns = [f'col_{i}' for i in range(df_filtered.shape[1])]
# UWAGA!! WYBIERZ KOLUMNY KT CIE INTERESUJA
pattern_columns = ['col_0', 'col_1', 'col_12']
df_relevant = df_filtered[pattern_columns]

df_relevant.loc[:, 'col_0'] = df_relevant['col_0'].ffill()  # Data
df_relevant.loc[:, 'col_1'] = df_relevant['col_1'].ffill()  # Godzina

df_relevant.columns = ['Data', 'Godzina', 'Opis']
df_relevant = df_relevant.dropna(subset=['Opis', 'Data', 'Godzina']).reset_index(drop=True)

df_relevant['Data'] = pd.to_datetime(df_relevant['Data'], errors='coerce')


def process_time_range(time_range):
    try:
        if time_range in ['sobota', 'niedziela']:
            return None, None
        start_time_str, end_time_str = time_range.split('-')
        start_time = datetime.strptime(start_time_str.strip(), '%H:%M').time()
        end_time = datetime.strptime(end_time_str.strip(), '%H:%M').time()

        start_time = (datetime.combine(datetime.today(), start_time) - timedelta(hours=1)).time()
        end_time = (datetime.combine(datetime.today(), end_time) - timedelta(hours=1)).time()
        return start_time, end_time
    except Exception as e:
        print(f"Niepoprawny format godzin: {time_range}. Błąd: {e}")
        return None, None


df_relevant[['Godzina_Start', 'Godzina_Koniec']] = df_relevant['Godzina'].apply(
    lambda x: pd.Series(process_time_range(str(x))))

df_relevant = df_relevant.dropna(subset=['Godzina_Start', 'Godzina_Koniec']).reset_index(drop=True)

title_dict = {
    "OB": "dr Olaf Bar",
    "MB": "dr hab. inż. Michał Bereta, prof. PK",
    "JB": "dr inż. Jerzy Białas",
    "LB": "dr hab. inż. Lesław Bieniasz, prof. PK",
    "PB": "mgr inż. Piotr Biskup",
    "GB": "mgr inż. Grzegorz Bogdał",
    "BB": "dr Barbara Borowik",
    "TB": "prof. zw.dr hab. inż. Tadeusz Burczyński",
    "DC": "Dominika Cywicka",
    "SD": "prof.dr hab. Stanisław Drożdż",
    "PD": "dr Piotr Drygaś",
    "SF": "dr hab.inż. Sergiy Fialko, prof. PK",
    "MG": "mgr inż. Michał Gandor",
    "ŁG": "mgr inż. Łukasz Gaża",
    "TG": "dr inż. Tomasz Gąciarz",
    "DG": "dr inż. Daniel Grzonka",
    "AJ": "dr Agnieszka Jakóbik",
    "LJ": "dr inż. Lech Jamroż",
    "PJ": "dr inż. Paweł Jarosz",
    "AJ-S": "dr inż. Anna Jasińska-Suwada",
    "MJ": "dr hab. inż. Maciej Jaworski, prof. PK",
    "JK": "dr hab. Joanna Kołodziej, prof. PK",
    "FK": "dr inż. Filip Krużel",
    "WK": "mgr inż. Wojciech Książek",
    "DK": "mgr inż. Dominik Kulis",
    "RK": "dr Radosław Kycia",
    "JL": "dr hab. inż. Jacek Leśkow, prof. PK",
    "PŁ": "dr inż. Piotr Łabędż",
    "ML": "mgr inż. Maryna Łukaczyk",
    "PM": "prof. dr hab. inż. Piotr Malecki",
    "AM": "dr Adam Marszałek",
    "MM": "mgr inż. Mateusz Michałek",
    "AM": "mgr inż. Andrzej Mycek",
    "MNaw": "mgr inż. Mateusz Nawrocki",
    "MNied": "mgr inż. Michał Niedźwiecki",
    "AN": "mgr inż. Artur Niewiarowski",
    "MN": "mgr inż. Mateusz Nytko",
    "JO": "mgr inż. Jerzy Orlof",
    "PO": "dr hab. inż. arch. Paweł Ozimek, prof. PK",
    "AP": "dr inż. Anna Plichta",
    "PPł": "dr hab. inż. Paweł Pławiak, prof. PK",
    "JPł": "dr inż. Joanna Płażek",
    "WR": "prof.dr hab. inż. Waldemar Rachowicz",
    "ARP": "Aleksander Radwan-Pragłowski",
    "DR": "mgr inż. Dominika Rola",
    "MR": "mgr inż. Mirosław Roszkowski",
    "KS": "dr inż. Krzysztof Skabek",
    "MS": "dr hab. inż. Marek Stanuszek, prof. PK",
    "PSz": "mgr inż. Piotr Szuster",
    "KSw": "mgr inż. Krzysztof Swałdek",
    "IU": "dr Ilona Urbaniak",
    "MW": "dr Marcin Wątorek",
    "AWid": "mgr inż. Adrian Widłak",
    "AW": "dr inż. Andrzej Wilczyński",
    "JW.": "mgr inż. Jan Wojtas",
    "AWoź": "mgr inż. Andrzej Woźniacki",
    "AWyrz": "dr Adam Wyrzykowski",
    "JZ": "dr inż. Jerzy Zaczek",
    "Mzom": "dr Maryam Zomorodi Moghaddam, prof. PK",
    "DŻ": "dr inż. Dariusz Żelasko"
}
def add_titles_to_comment(opis):
    words = opis.split()
    titles = [title_dict.get(word, word) for word in words if word in title_dict]
    return ", ".join(titles)  # dodaj tytuły wykl. do komentarza

def remove_initials_from_description(description):
    parts = description.split()
    filtered_parts = [parts[0]] + [part for part in parts[1:] if not (part.isupper() and 2 <= len(part) <= 3)]
    return ' '.join(filtered_parts)

df_relevant['Komentarz'] = df_relevant['Opis'].apply(add_titles_to_comment)
df_relevant['Opis'] = df_relevant['Opis'].apply(remove_initials_from_description)

def create_ics(dataframe, filename="plan_zajec.ics"):
    calendar = Calendar()
    if dataframe.empty:
        print("Brak danych do dodania do kalendarza.")
        return
    for _, row in dataframe.iterrows():
        event = Event()
        event.name = row["Opis"]
        event.begin = f"{row['Data'].strftime('%Y-%m-%d')} {row['Godzina_Start']}"
        event.end = f"{row['Data'].strftime('%Y-%m-%d')} {row['Godzina_Koniec']}"

        comment = row['Komentarz']
        if comment:
            event.description = comment

        if "zdalnie" not in row["Opis"].lower() and "wykład" not in row["Opis"].lower():
            event.location = "Warszawska 24, Kraków, Politechnika Krakowska"

        calendar.events.add(event)

    with open(filename, "w", encoding="utf-8") as f:
        f.writelines(calendar)


desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "plan_zajec.ics")
create_ics(df_relevant, filename=desktop_path)

print(f"Plik .ics został zapisany na pulpicie: {desktop_path}")
