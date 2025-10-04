# ExcelPixelArt

Aplikacja Streamlit do zamiany obrazów na mozaikę pikseli oraz generowania arkuszy Excel z gotowym układem komórek.

Dostępna na https://excel40lat.streamlit.app/

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

## Jak działa algorytm – krok po kroku, po ludzku

### 1. Przygotowanie zdjęcia

1. **Wczytanie obrazu** – aplikacja otwiera plik i konwertuje go do formatu RGBA, czyli z informacją o przezroczystości. Dzięki temu łatwiej zdecydować, które piksele są tłem, a które właściwym zdjęciem.
2. **Usuwanie tła** – w zależności od wybranego trybu program:
   - zostawia przezroczyste fragmenty z oryginalnego PNG (tryb `alpha`),
   - automatycznie pobiera próbkę koloru z rogów i usuwa podobne odcienie (tryb `auto-corners`),
   - korzysta z ręcznie wskazanego koloru (tryb `manual-color`).
   W praktyce oznacza to, że z portretu znika jednolite tło, a komórki Excela pozostają czyste.

### 2. Uproszczenie zdjęcia do mozaiki

1. **Zmniejszenie szerokości** – obraz jest przeskalowywany do liczby kolumn zadeklarowanej w panelu bocznym. Wysokość dobierana jest proporcjonalnie, żeby twarz lub obiekt nie został zniekształcony.
2. **Zamiana kolorów na krótką paletę** – z wykorzystaniem funkcji „Adaptive Palette” z biblioteki Pillow, aplikacja wyszukuje najczęściej spotykane kolory i ogranicza ich liczbę do wskazanej wartości (np. 32 lub 64).
3. **Zapisanie siatki pikseli** – dla każdego miejsca w mozaice zapamiętywany jest kolor albo informacja „przezroczyste” (gdy piksel należy do tła). To ta siatka steruje później wypełnianiem komórek w Excelu.

### 3. Generowanie banera z tekstem

1. **Stworzenie tła** – powstaje prostokąt w wybranym kolorze, który ma tyle kolumn, co mozaika, oraz liczbę wierszy ustawioną w panelu.
2. **Delikatna tekstura** – co kilka kolumn i wierszy nanoszone są jaśniejsze linie. Tworzą one wrażenie arkusza Excela, ale nie rozmywają liter, bo zmiana koloru jest subtelna.
3. **Specjalna „pikselowa” czcionka** – litery i cyfry są rysowane na bazie gotowych szablonów 5×7 punktów. Program skaluje je tylko całkowitą liczbę razy (2×, 3× itd.), dzięki czemu krawędzie są ostre jak w retro grach.
4. **Normalizacja napisów** – polskie znaki są zamieniane na ich uproszczone odpowiedniki (`Ł` → `L` itd.). Gdyby w czcionce zabrakło znaku diakrytycznego, tekst wyglądałby dziwnie; ta zamiana gwarantuje spójny wygląd.
5. **Dobór wielkości liter** – algorytm oblicza, ile miejsca zajmują litery przy danym powiększeniu. Jeśli napis się nie mieści, automatycznie zmniejsza skalę, zachowując minimalny rozmiar zapewniający czytelność.
6. **Rozmieszczenie** – nagłówek jest centrowany i umieszczany w górnej części banera, podtytuł trafia niżej. Dzięki temu napisy są wyśrodkowane i mają odpowiedni oddech.

### 4. Budowa pliku Excel

1. **Przygotowanie arkusza** – wyłączane są domyślne linie siatki, a kolumny i wiersze ustawiane tak, by tworzyły niemal kwadratowe komórki.
2. **Malowanie mozaiki i banera** – każdy kolor z siatki pikseli zamienia się w wypełnienie komórki (`PatternFill`). Jeśli piksel był przezroczysty, komórka zostaje pusta (lub przyjmuje kolor tła, jeśli użytkownik włączy tę opcję).
3. **Łączenie elementów** – mozaika i baner mogą być ułożone jeden pod drugim albo obok siebie. Między nimi można dodać przerwę dla lepszej czytelności.

### Dlaczego litery na banerze są ostre i czytelne?

- **Pikselowe rysowanie znaków** – każda litera to ręcznie zdefiniowany wzór 5×7. Program wypełnia konkretne „kwadraciki”, a nie próbuje drukować fontów wektorowych. Dzięki temu w Excelu widzimy ostre prostokąty, a nie rozmyte kształty.
- **Skalowanie bez rozmycia** – wzory są powiększane tylko całkowitym mnożnikiem. Brak pośrednich wartości (np. 2,4×) eliminuje efekt rozmytych krawędzi.
- **Kontrola rozmiaru** – nagłówek nie spada poniżej ustawionej minimalnej skali, żeby zachować czytelność nawet przy węższych banerach. Jeśli napis jest za długi, program zmniejsza go stopniowo, aż idealnie zmieści się w szerokości.
- **Upraszczanie znaków specjalnych** – diakrytyki i myślniki są zamieniane na wersje zgodne z przygotowanym alfabetem. To zapobiega dziurom w napisie (np. brakującemu „Ś”).
- **Delikatny kontrast** – tło może mieć subtelną fakturę, ale różnica jasności jest niewielka. Litery nadal mają pełne, jednolite wypełnienie, więc pozostają dobrze odczytywalne.

## Weryfikacja zmian

Projekt nie posiada testów automatycznych. Po modyfikacjach:

- uruchom aplikację,
- wgraj przykładowy obraz,
- sprawdź podgląd mozaiki i banera,
- pobierz wygenerowany arkusz i otwórz go w Excelu dla szybkiego sprawdzenia
