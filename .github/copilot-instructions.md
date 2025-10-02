# Przewodnik AI dla ExcelPixelArt

- **Szybki start**
  - Repozytorium zawiera jedną aplikację Streamlit (`streamlit_excel_pixel_art_app.py`). Uruchamiaj ją komendą `streamlit run streamlit_excel_pixel_art_app.py`.
  - Zależności są w `requirements.txt` (`streamlit`, `pillow`, `openpyxl`). Przed startem środowiska wykonaj `pip install -r requirements.txt`.
  - Brak testów automatycznych. Po zmianach zawsze wgraj przykładowy obraz, sprawdź podglądy w Streamlit i pobierz `.xlsx`, aby zrobić szybki smoke test.

- **Główna logika** (`streamlit_excel_pixel_art_app.py`)
  - Suwaki w panelu bocznym konfigurują `MosaicOptions`. Funkcja `image_to_pixel_grid` skaluje obraz, usuwa tło (alpha / rogi / kolor ręczny), kwantyzuje barwy w Pillow i zwraca siatkę `List[List[Pixel]]`, gdzie `None` oznacza „nie koloruj komórki w Excelu”.
  - Baner generuje `banner_to_pixel_grid` na podstawie `BannerOptions`. Renderer korzysta z bitmapowej czcionki 5×7 (`FONT_5x7`, `normalize_ascii`, `blit_text`) oraz helperów `lighten`/`darken`, aby dodać teksturę tabeli i poprawnie obsłużyć polskie znaki (mapowane do ASCII). Przy dodawaniu nowego tekstu zachowaj logikę zmniejszania skali, dopóki napis mieści się w dostępnych kolumnach.
  - Podglądy buduje funkcja `grid_to_image`, następnie obrazy są powiększane `Image.NEAREST`, aby zachować ostre piksele. Każda nowa siatka powinna trafiać równocześnie do podglądu i do eksportu.
  - Excel generuje tylko `build_workbook`. Funkcja pilnuje kwadratowych komórek (`ensure_square_cells`), używa `FillCache` do buforowania kolorów i honoruje przełącznik `layout` (pionowy vs. obok siebie). Zachowaj konwencję `None` → „pomiń wypełnienie”, aby nie psuć przezroczystości.

- **Konwencje i pułapki**
  - Siatki pikseli traktuj jak prostokątne, nierozszerzalne listy. Każdy wiersz musi mieć identyczną długość zanim przekażesz dane dalej.
  - Stałe kolorów znajdują się na górze pliku (`EXCEL_GREEN`, `WHITE`, `BLACK`). Jeśli dodajesz nowe motywy, zamieniaj RGB na format ARGB funkcją `to_hex_argb`, bo `FillCache` używa go jako klucza.
  - UI jest po polsku — trzymaj się tego stylu w nowych labelach, tooltipach i przyciskach.
  - Domyślne wartości `MosaicOptions.target_width` i `palette_colors` odpowiadają maksymalnym wartościom suwaków (220 / 64). Tworząc nowe kontrolki, zadbaj o spójne domyślne ustawienia w dataclassie i w UI, by eksport odpowiadał podglądowi.
  - W całej aplikacji używamy `st.image(..., use_container_width=True)`. Jeśli Streamlit zastąpi ten parametr, zmień wszystkie wystąpienia jednocześnie, żeby uniknąć niespójnego layoutu.

- **Rozszerzanie aplikacji**
  - Przy nowych formatach eksportu ponownie wykorzystuj siatki z `image_to_pixel_grid` i `banner_to_pixel_grid`, zamiast liczyć piksele od zera.
  - Dodając style banera, wprowadzaj warunki w `banner_to_pixel_grid` i trzymaj powiązanie rozmiaru tekstu z liczbą wierszy (`rows`) tak jak w aktualnej implementacji.
  - Jeżeli wprowadzisz testy lub narzędzia CLI, umieszczaj je obok głównego skryptu i korzystaj z tych samych helperów, żeby nie duplikować logiki obrazu i Excela.
