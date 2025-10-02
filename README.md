# ExcelPixelArt

Aplikacja Streamlit do zamiany obrazów na mozaikę pikseli oraz generowania arkuszy Excel z gotowym układem komórek.

## Szybki start

1. Zainstaluj zależności:
   ```powershell
   pip install -r requirements.txt
   ```
2. Uruchom aplikację:
   ```powershell
   streamlit run streamlit_excel_pixel_art_app.py
   ```
3. W przeglądarce wgraj obraz (najlepiej PNG z przezroczystością), dostosuj ustawienia w panelu bocznym i pobierz wygenerowany plik `.xlsx`.

## Kluczowe funkcje

- `streamlit_excel_pixel_art_app.py` – główny skrypt z logiką przekształcania obrazu, generowania banera i eksportu do Excela.
- `requirements.txt` – lista bibliotek: Streamlit (UI), Pillow (obróbka grafiki) oraz openpyxl (tworzenie arkusza Excel).

## Weryfikacja zmian

Projekt nie posiada testów automatycznych. Po modyfikacjach:

- uruchom aplikację,
- wgraj przykładowy obraz,
- sprawdź podgląd mozaiki i banera,
- pobierz wygenerowany arkusz i otwórz go w Excelu dla szybkiego sprawdzenia
