import asyncio
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime
import os
import psutil
import subprocess
import time

class OzonPvzBot:
    def __init__(self):
        self.browser = None
        self.page = None
        self.playwright = None
    
    def find_opened_edge_windows(self):
        """Поиск открытых окон Edge с turbo-pvz.ozon.ru"""
        print("🔍 Ищу открытые окна Microsoft Edge...")
        
        try:
            # Получаем список всех процессов Edge
            edge_processes = []
            for process in psutil.process_iter(['pid', 'name', 'cmdline']):
                try:
                    if process.info['name'] and 'msedge' in process.info['name'].lower():
                        edge_processes.append(process.info)
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            if not edge_processes:
                print("❌ Не найдено открытых процессов Edge")
                return False
            
            print(f"✅ Найдено процессов Edge: {len(edge_processes)}")
            
            # Проверяем, есть ли окно с turbo-pvz.ozon.ru
            for process in edge_processes:
                if process['cmdline']:
                    cmdline = ' '.join(process['cmdline'])
                    if 'turbo-pvz.ozon.ru' in cmdline:
                        print("✅ Найдено окно с turbo-pvz.ozon.ru")
                        return True
            
            print("ℹ️  Окно с turbo-pvz.ozon.ru не найдено, но Edge запущен")
            return True
            
        except Exception as e:
            print(f"❌ Ошибка при поиске процессов: {e}")
            return False
    
    async def connect_to_existing_edge(self):
        """Подключение к существующему окну Edge"""
        print("🔗 Пытаюсь подключиться к существующему Edge...")
        
        self.playwright = await async_playwright().start()
        
        # Способ 1: Попробуем подключиться через CDP
        try:
            # Стандартный порт для отладки Edge
            self.browser = await self.playwright.chromium.connect_over_cdp("http://localhost:9222")
            contexts = self.browser.contexts
            
            if contexts:
                for context in contexts:
                    pages = context.pages
                    if pages:
                        for page in pages:
                            if "turbo-pvz.ozon.ru" in page.url:
                                self.page = page
                                print("✅ Подключился к существующей вкладке turbo-pvz.ozon.ru")
                                return True
                
                # Если не нашли нужную вкладку, используем первую доступную
                self.page = contexts[0].pages[0] if contexts[0].pages else await contexts[0].new_page()
                print("✅ Использую первую доступную вкладку")
                return True
                
        except Exception as e:
            print(f"❌ Не удалось подключиться через CDP: {e}")
        
        # Способ 2: Запускаем Edge с отладочным портом
        print("🔄 Пробую запустить Edge с отладочным портом...")
        try:
            # Закрываем все предыдущие процессы Edge (осторожно!)
            subprocess.run(['taskkill', '/f', '/im', 'msedge.exe'], capture_output=True)
            time.sleep(2)
            
            # Запускаем Edge с отладочным портом
            edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            if not os.path.exists(edge_path):
                edge_path = r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            
            subprocess.Popen([
                edge_path,
                '--remote-debugging-port=9222',
                '--user-data-dir=./edge_profile',
                'https://turbo-pvz.ozon.ru/'
            ])
            
            time.sleep(5)  # Ждем запуска
            
            # Подключаемся к новому экземпляру
            self.browser = await self.playwright.chromium.connect_over_cdp("http://localhost:9222")
            self.page = self.browser.contexts[0].pages[0] if self.browser.contexts[0].pages else await self.browser.contexts[0].new_page()
            
            print("✅ Запустил Edge с отладкой и подключился")
            return True
            
        except Exception as e:
            print(f"❌ Не удалось запустить Edge с отладкой: {e}")
        
        # Способ 3: Ручное подключение
        print("👆 Ручной режим подключения")
        print("Пожалуйста, выполните следующие шаги:")
        print("1. Убедитесь, что Microsoft Edge открыт")
        print("2. Вы должны быть на странице turbo-pvz.ozon.ru")
        print("3. Убедитесь, что вы авторизованы")
        input("Нажмите Enter после выполнения этих шагов...")
        
        try:
            # Пробуем найти существующий браузер
            self.browser = await self.playwright.chromium.launch(channel="msedge", headless=False)
            contexts = self.browser.contexts
            
            if contexts and contexts[0].pages:
                self.page = contexts[0].pages[0]
                print("✅ Нашел открытое окно Edge")
                return True
            else:
                # Создаем новую вкладку в существующем браузере
                self.page = await self.browser.new_page()
                await self.page.goto("https://turbo-pvz.ozon.ru/")
                print("✅ Создал новую вкладку с сайтом")
                return True
                
        except Exception as e:
            print(f"❌ Критическая ошибка: {e}")
            return False
    
    async def ensure_correct_page(self):
        """Убедиться, что мы на правильной странице"""
        current_url = self.page.url
        
        if "turbo-pvz.ozon.ru" not in current_url:
            print("🌐 Перехожу на turbo-pvz.ozon.ru...")
            await self.page.goto("https://turbo-pvz.ozon.ru/", wait_until="networkidle")
            await self.page.wait_for_timeout(1000)
        
        # Проверяем авторизацию
        if "auth" in self.page.url.lower() or "login" in self.page.url.lower():
            print("❌ Требуется авторизация!")
            print("Пожалуйста, войдите в аккаунт в открытом браузере")
            input("Нажмите Enter после авторизации...")
            await self.page.goto("https://turbo-pvz.ozon.ru/", wait_until="networkidle")
        
        print("✅ Страница готова к работе")
        return True
    
    #####################################################################################################################
    
    async def navigate_to_shipment_transport(self):
        """Переход в раздел Отправка -> Отправка перевозок"""
        print("📂 Ищу раздел 'Отправка'...")
        
        await self.page.wait_for_timeout(2000)
        
        # Проверяем, не находимся ли уже в нужном разделе
        current_url = self.page.url
    #    if "outbound" in current_url and ("transport" in current_url or "id=-1000" in current_url):
    #        print("✅ Уже в разделе отправки перевозок")
    #       return True
        
        # Если мы в общем разделе отправки, но не в перевозках, ищем кнопку перевозок
        if "outbound" in current_url and "id=-1000" in current_url:
            print("ℹ️  В разделе отправки, но не в перевозках")
            return await self.select_shipment_transport()
        
        # Шаг 1: Нажимаем на "Отправка" в меню
        shipment_selectors = [
            "a[href*='outbound']",
            "text=Отправка",
            "a:has-text('Отправка')", 
            "li:has-text('Отправка')",
            "div:has-text('Отправка')",
            "[class*='shipment']",
        ]
        
        for selector in shipment_selectors:
            try:
                await self.page.click(selector, timeout=5000)
                print(f"✅ Нажал на 'Отправка': {selector}")
                await self.page.wait_for_timeout(3000)
                
                # После клика на Отправка, сразу ищем перевозки
                return await self.select_shipment_transport()
                        
            except Exception as e:
                continue
        
        print("❌ Раздел 'Отправка' не найден автоматически")
        return False

    async def select_shipment_transport(self):
        """Выбор подпункта 'Отправка перевозок' после входа в Отправка"""
        print("🚚 Ищу подпункт 'Отправка перевозок'...")
        
        await self.page.wait_for_timeout(1000)
        
        transport_selectors = [
            "text=Отправка перевозок",
            "text=перевозок",
            "a:has-text('перевозок')",
            "div:has-text('перевозок')",
            "button:has-text('перевозок')",
            "[class*='transport']",
            "[class*='перевозок']",
            "a[href*=id=-1001]",
        ]
        
        for selector in transport_selectors:
            try:
                await self.page.click(selector, timeout=5000)
                print(f"✅ Нажал на 'Отправка перевозок': {selector}")
                await self.page.wait_for_timeout(1000)
                return True
            except Exception as e:
                continue
        
        print("❌ Подпункт 'Отправка перевозок' не найден автоматически")
        return False

    async def select_flow_type(self, flow_type):
        """Выбор типа потока"""
        print(f"🔘 Ищу тип потока: {flow_type}...")
        
        await self.page.wait_for_timeout(2000)
        
        # Ищем все элементы с типами потоков
        flow_elements = await self.page.query_selector_all("div._flowType_16tx4_203")
        
        for element in flow_elements:
            try:
                text = await element.text_content()
                if flow_type in text:
                    await element.click()
                    print(f"✅ Выбрал: {flow_type}")
                    await self.page.wait_for_timeout(2000)
                    return True
            except Exception as e:
                continue
        
        print(f"❌ Тип потока '{flow_type}' не найден автоматически")
        return False

    async def find_target_block_with_informer(self):
        """Поиск блока с информатором 'Добавьте содержимое в перевозку'"""
        print("🔍 Ищу блок с информатором 'Добавьте содержимое в перевозку'...")
        
        await self.page.wait_for_timeout(3000)
        
        # Ищем все блоки с классом _block_4j0aa_1
        blocks = await self.page.query_selector_all("div._block_4j0aa_1")
        print(f"✅ Найдено блоков _block_4j0aa_1: {len(blocks)}")
        
        for i, block in enumerate(blocks):
            try:
                # Ищем информатор внутри блока
                informer_text = await block.inner_text()
                if "Добавьте содержимое в перевозку" in informer_text:
                    print(f"✅ Найден целевой блок (блок #{i + 1}) с нужным информатором")
                    return block
            except Exception as e:
                continue
        
        print("❌ Не найден блок с информатором 'Добавьте содержимое в перевозку'")
        return None

    async def extract_items_before_locked_section(self, target_block):
        """Извлечение товаров, которые находятся ДО секции 'Не подходит направление потока'"""
        print("📦 Ищу товары ДО секции 'Не подходит направление потока'...")
        
        items_data = []
        
        try:
            # Находим все элементы товаров в блоке
            all_elements = await target_block.query_selector_all("div._element_16tx4_1")
            print(f"✅ Найдено всех элементов товаров: {len(all_elements)}")
            
            # Парсим только незаблокированные товары
            for i, element in enumerate(all_elements):
                try:
                    # Проверяем, заблокирован ли элемент
                    is_locked = await element.query_selector("div._locked_16tx4_13") is not None
                    
                    if not is_locked:
                        item_data = await self.parse_item(element, i, "требуется оформление")
                        if item_data:
                            items_data.append(item_data)
                            print(f"✅ Добавлен товар: {item_data['Название'][:50]}...")
                    
                except Exception as e:
                    print(f"❌ Ошибка парсинга элемента {i}: {e}")
                    continue
            
            print(f"✅ Найдено незаблокированных товаров: {len(items_data)}")
            
        except Exception as e:
            print(f"❌ Ошибка при извлечении товаров: {e}")
        
        return items_data

    async def parse_item(self, item_element, index, status):
        """Парсинг карточки товара"""
        try:
            # Название товара
            name_element = await item_element.query_selector("div._titleWrap_1ailj_14")
            name = await name_element.text_content() if name_element else "Неизвестно"
            
            # Штрихкод
            barcode_element = await item_element.query_selector("div._barcode_1ailj_7")
            barcode_text = await barcode_element.text_content() if barcode_element else ""
            barcode = barcode_text.replace("Штрихкод:", "").strip() if barcode_text else "Не найден"
            
            # Местоположение (ячейка)
            location_element = await item_element.query_selector("div._address_1ailj_28 .ozi__badge__label__Rb41r")
            location = await location_element.text_content() if location_element else "Не указано"
            
            # Тип потока
            flow_element = await item_element.query_selector("div._flowType_16tx4_203")
            flow_type = await flow_element.text_content() if flow_element else "Не определен"
            
            # Цвет (если есть)
            color_element = await item_element.query_selector("div._flex_lxoww_1 div:last-child")
            color = await color_element.text_content() if color_element else ""
            
            # Размер (если есть)
            size = ""
            content_element = await item_element.query_selector("div._content_1ailj_36")
            if content_element:
                content_text = await content_element.text_content()
                if "Размер" in content_text:
                    size_parts = content_text.split("Размер")
                    if len(size_parts) > 1:
                        size = size_parts[1].strip().split('\n')[0]
            
            # Статус блокировки
            is_locked = await item_element.query_selector("div._locked_16tx4_13") is not None
            lock_status = "Заблокирован" if is_locked else "Доступен"
            
            return {
                '№': index + 1,
                'Название': name.strip(),
                'Штрихкод': barcode,
                'Ячейка': location,
                'Тип_потока': flow_type,
                'Цвет': color.strip(),
                'Размер': size.strip(),
                'Статус_блокировки': lock_status,
                'Статус_оформления': status
            }
        except Exception as e:
            print(f"❌ Ошибка парсинга элемента {index}: {e}")
            return None

    async def run(self, flow_type="Прямой поток"):
        """Основной метод"""
        try:
            if not await self.connect_to_existing_edge():
                print("❌ Не удалось подключиться к Edge")
                return
            
            print(f"🎯 Начинаю сбор данных для: {flow_type}")
            print("📋 Этапы работы:")
            print("1. Переход в 'Отправка'")
            print("2. Выбор 'Отправка перевозок'")
            print("3. Выбор типа потока")
            print("4. Сбор товаров ДО секции 'Не подходит направление потока'")
            
            # Убеждаемся, что на правильной странице
            if not await self.ensure_correct_page():
                print("❌ Не удалось перейти на нужную страницу")
                return
            
            # Шаг 1: Переход в Отправка -> Отправка перевозок
            if not await self.navigate_to_shipment_transport():
                print("❌ Не удалось перейти в раздел отправки перевозок")
                
                # Пробуем прямой переход
                print("🔄 Пробую прямой переход на страницу перевозок...")
                await self.page.goto("https://turbo-pvz.ozon.ru/outbound?id=-1000", wait_until="networkidle")
                await self.page.wait_for_timeout(5000)
                
                # Проверяем, загрузилась ли страница
                current_url = self.page.url
                print(f"🌐 Текущий URL после прямого перехода: {current_url}")
                
                if "transport" not in current_url and "id=-1000" not in current_url:
                    print("❌ Прямой переход тоже не сработал")
                    return
            
            # Проверяем, что мы действительно на странице перевозок
            current_url = self.page.url
            print(f"🌐 Текущий URL: {current_url}")
            
            # Шаг 2: Выбор типа потока
            if not await self.select_flow_type(flow_type):
                print(f"❌ Не удалось выбрать тип потока '{flow_type}'")
                return
            
            # Шаг 3: Находим целевой блок с информатором
            target_block = await self.find_target_block_with_informer()
            if not target_block:
                print("❌ Не найден целевой блок")
                return
            
            # Шаг 4: Извлекаем товары ДО секции "Не подходит направление потока"
            print("⏳ Загружаю данные о товарах ДО секции 'Не подходит направление потока'...")
            items_data = await self.extract_items_before_locked_section(target_block)
            
            if not items_data:
                print("❌ Не найдено товаров ДО заблокированной секции")
                return
            
            self.display_results(items_data, flow_type)
            self.save_to_files(items_data, flow_type)
            
        except Exception as e:
            print(f"❌ Ошибка: {e}")
        finally:
            await self.close()
            print("✅ Работа завершена")

    def display_results(self, items_data, flow_type):
        """Вывод результатов"""
        print(f"\n{'='*80}")
        print(f"📊 РЕЗУЛЬТАТЫ: {flow_type}")
        print(f"👉 Товаров ДО секции 'Не подходит направление потока': {len(items_data)}")
        print(f"{'='*80}")
        
        for item in items_data:
            print(f"\n{item['№']}. {item['Название']}")
            print(f"   📍 Ячейка: {item['Ячейка']}")
            print(f"   🏷️  Штрихкод: {item['Штрихкод']}")
            print(f"   🔄 Тип потока: {item['Тип_потока']}")
            if item['Цвет']:
                print(f"   🎨 Цвет: {item['Цвет']}")
            if item['Размер']:
                print(f"   📏 Размер: {item['Размер']}")
            print(f"   🔒 Статус блокировки: {item['Статус_блокировки']}")

    def save_to_files(self, items_data, flow_type):
        """Сохранение в файлы"""
        if not items_data:
            return
        
        os.makedirs("results", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_filename = f"results/ozon_shipment_{flow_type.replace(' ', '_')}_{timestamp}"
        
        # Excel
        df = pd.DataFrame(items_data)
        excel_filename = f"{base_filename}.xlsx"
        df.to_excel(excel_filename, index=False)
        print(f"💾 Excel файл сохранен: {excel_filename}")
        
        # TXT
        txt_filename = f"{base_filename}.txt"
        with open(txt_filename, 'w', encoding='utf-8') as f:
            f.write(f"ОТЧЕТ: {flow_type}\n")
            f.write(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write(f"Товаров ДО секции 'Не подходит направление потока': {len(items_data)}\n")
            f.write("=" * 60 + "\n\n")
            
            for i, item in enumerate(items_data, 1):
                f.write(f"ТОВАР #{i}\n")
                f.write(f"Название: {item['Название']}\n")
                f.write(f"Ячейка: {item['Ячейка']}\n")
                f.write(f"Штрихкод: {item['Штрихкод']}\n")
                f.write(f"Тип потока: {item['Тип_потока']}\n")
                if item['Цвет']:
                    f.write(f"Цвет: {item['Цвет']}\n")
                if item['Размер']:
                    f.write(f"Размер: {item['Размер']}\n")
                f.write(f"Статус блокировки: {item['Статус_блокировки']}\n")
                f.write("-" * 40 + "\n\n")

    async def close(self):
        """Корректное закрытие ресурсов"""
        try:
            if self.browser:
                await self.browser.close()
            if self.playwright:
                await self.playwright.stop()
        except Exception as e:
            print(f"⚠️  Предупреждение при закрытии: {e}")

async def main():
    """Основная функция"""
    print("🤖 Бот для Турбо ПВЗ")
    print("👉 Собирает товары ДО секции 'Не подходит направление потока'")
    print("=" * 50)
    
    print("Типы потоков:")
    print("1 - Прямой поток")
    print("2 - Возвратный поток")
    
    choice = input("Выберите тип (1 или 2): ").strip()
    
    if choice == "2":
        flow_type = "Возвратный поток"
    else:
        flow_type = "Прямой поток"
    
    bot = OzonPvzBot()
    await bot.run(flow_type)

if __name__ == "__main__":
    print("Установите: pip install playwright pandas openpyxl psutil")
    input("Нажмите Enter для запуска...")
    
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n🚫 Программа прервана пользователем")
    except Exception as e:
        print(f"💥 Критическая ошибка: {e}")
