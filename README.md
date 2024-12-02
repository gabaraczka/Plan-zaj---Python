# Plan-zajęć-Python

**Ten skrypt umożliwia przetwarzanie harmonogramu zajęć z pliku Excel i generowanie pliku .ics kompatybilnego z kalendarzami Google i Apple.** 

Wymaganie zainstalowania biblioteki:

  **pandas**
  
  **ics**

Skopiuj kod

pip install pandas ics


Jak używać?

Skonfiguruj plik wejściowy:

Wskaż pełną ścieżkę do pliku Excel w miejscu <<< WSTAW ŚCIEŻKĘ >>>.
Zmień nazwę arkusza w miejscu <<< WSTAW NAZWĘ ARKUSZA >>> na odpowiednią nazwę arkusza w pliku Excel.


Dostosuj kolumny:

Skrypt zakłada, że w pliku Excel znajdują się odpowiednie kolumny z datą, godzinami i opisem zajęć. Zmień zmienną pattern_columns, aby wskazać odpowiednie kolumny w pliku, które zawierają dane, które chcesz przetworzyć.

Skrypt zapisuje plik .ics na pulpicie w formacie kompatybilnym z kalendarzami Google/Apple.
