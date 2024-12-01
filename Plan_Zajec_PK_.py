import pandas as pd
from ics import Calendar, Event
import os

from datetime import datetime

# Ścieżka do pliku wejściowego
file_path = r'<<< WSTAW ŚCIEŻKE!!!>>>>'
df = pd.read_excel(file_path, sheet_name='<<< WSTAW NAZWE ARKUSZA >>>>')

df_filtered = df.iloc[5:].reset_index(drop=True)

df_filtered.columns = [f'col_{i}' for i in range(df_filtered.shape[1])]


##### PRZED ROZPOCZECIEM PRACY ZMIEN KOLUMNY KT CIE DOTYCZA
pattern_columns = ['col_0', 'col_1', 'col_12']
df_relevant = df_filtered[pattern_columns]


df_relevant.loc[:, 'col_0'] = df_relevant['col_0'].ffill()  # Data
df_relevant.loc[:, 'col_1'] = df_relevant['col_1'].ffill()  # Godzina

df_relevant.columns = ['Data', 'Godzina', 'Opis']
df_relevant = df_relevant.dropna(subset=['Opis', 'Data', 'Godzina']).reset_index(drop=True)


df_relevant['Data'] = pd.to_datetime(df_relevant['Data'], errors='coerce')


# funkcja rozdzielająca godziny na początek i koniec
def process_time_range(time_range):
    try:
        #  "8:00-10:30"
        if time_range == 'sobota' or time_range == 'niedziela':
            pass
        else:
            start_time_str, end_time_str = time_range.split('-')

            start_time = datetime.strptime(start_time_str.strip(), '%H:%M').time()
            end_time = datetime.strptime(end_time_str.strip(), '%H:%M').time()

            return start_time, end_time
    except Exception as e:
        print(f"Niepoprawny format godzin: {time_range}. Błąd: {e}")
        return None, None

df_relevant[['Godzina_Start', 'Godzina_Koniec']] = df_relevant['Godzina'].apply(
    lambda x: pd.Series(process_time_range(str(x))))

df_relevant = df_relevant.dropna(subset=['Godzina_Start', 'Godzina_Koniec']).reset_index(drop=True)

print(df_relevant.head())


# plik ics do kalendarza Google /apple

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

        calendar.events.add(event)

    with open(filename, "w", encoding="utf-8") as f:
        f.writelines(calendar)


desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "plan_zajec.ics")
create_ics(df_relevant, filename=desktop_path)

print(f"Plik .ics został zapisany na pulpicie: {desktop_path}")
