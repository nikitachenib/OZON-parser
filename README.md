# OZON-parser

1. Скачать python с [python.org](https://www.python.org/ftp/python/3.13.7/python-3.13.7-amd64.exe)
2. установить зависимости
   - pip install playwright pandas openpyxl psutil
3. уствновить браузеры через playwright
   - playwright install
4. для python3:
   - pip3 install playwright pandas openpyxl psutil
   - playwright install
5. Для Windows с py:
   - py -m pip install playwright pandas openpyxl psutil
   - py -m playwright install
6. (Linux/Mac):
   - pip install --user playwright pandas openpyxl psutil
   - playwright install
7. установка через requirements.txt
   - echo "playwright>=1.40.0" > requirements.txt
   - echo "pandas>=2.0.0" >> requirements.txt
   - echo "openpyxl>=3.1.0" >> requirements.txt
   - echo "psutil>=5.9.0" >> requirements.txt
  
   - pip install -r requirements.txt
8. выбрать путь до main.py
   - например: cd C:\OZON-parser\
9. ЗАПУСК:
    - python main.py
