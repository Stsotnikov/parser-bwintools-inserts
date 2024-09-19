import bs4
import xlsxwriter
import requests

# URL страницы для парсинга (общая часть)
main_url = 'https://russian.bwintools.com/'

# Заголовки для запроса (иногда требуется для обхода блокировок)
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36'
}

# Функция для получения и парсинга страницы
def get_soup(url):
    res = requests.get(url, headers=headers)
    res.encoding = 'utf-8'  # Принудительно устанавливаем кодировку
    return bs4.BeautifulSoup(res.text, 'html.parser')

# Функция для парсинга товара на странице товара
def parse_product_page(product_url):
    product_page = get_soup(product_url)

    # Извлекаем изображение 1
    image1_tag = product_page.select_one('#slidePic > ul > li.li.clickli.active > a > img')
    image1_url = image1_tag['src'] if image1_tag else '-'

    # Извлекаем изображение 2
    image2_tag = product_page.select_one('#slidePic > ul > li:nth-child(2) > a > img')
    image2_url = image2_tag['src'] if image2_tag else '-'

    # Извлекаем название товара
    title_tag = product_page.select_one('body > div.main-content.wrap-rule.fn-clear > div > div.chai_product_detailmain_lr > div > div.cont_r > h2')
    title = title_tag.get_text(strip=True) if title_tag else 'No Title'

    # Извлекаем Model Number Номер модели
    model_number_tag = product_page.select_one('body > div.main-content.wrap-rule.fn-clear > div > div.chai_product_detailmain_lr > div > div.cont_r > table:nth-child(3) > tbody > tr:nth-child(3) > td.p_attribute')
    model_number = model_number_tag.get_text(strip=True) if model_number_tag else 'No Title'

    # Извлекаем Product Name Название продукта
    product_name_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(2) > td:nth-child(2)')
    product_name = product_name_tag.get_text(strip=True) if product_name_tag else 'No Title'

    # Извлекаем Material Материал
    material_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(2) > td:nth-child(4)')
    material = material_tag.get_text(strip=True) if material_tag else 'No Title'

    # Извлекаем Usage Использование
    usage_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(3) > td:nth-child(4)')
    usage = usage_tag.get_text(strip=True) if usage_tag else 'No Title'

    # Извлекаем HRC твердость
    hrc_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(5) > td:nth-child(2)')
    hrc = hrc_tag.get_text(strip=True) if hrc_tag else 'No Title'

    # Извлекаем Application Применение
    application_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(6) > td:nth-child(2)')
    application = application_tag.get_text(strip=True) if application_tag else '-'

    # Извлекаем Высокий свет High Light 1-3
    high_light_tags = [product_page.select_one(f'#detail_infomation > table > tbody > tr:nth-child(7) > td > h2:nth-child({i})') for i in range(1, 4)]
    high_lights = [tag.get_text(strip=True) if tag else '-' for tag in high_light_tags]

    # Извлекаем Цвет / Color
    color_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(4) > td:nth-child(2)')
    color = color_tag.get_text(strip=True) if color_tag else '-'

    # Извлекаем Workpiece
    workpiece_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(3) > td:nth-child(2)')
    workpiece = workpiece_tag.get_text(strip=True) if workpiece_tag else '-'

    # Извлекаем Покрытие / coating
    coating_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(4) > td:nth-child(4)')
    coating = coating_tag.get_text(strip=True) if coating_tag else '-'

    # Извлекаем Особенность / Feature
    feature_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(5) > td:nth-child(4)')
    feature = feature_tag.get_text(strip=True) if feature_tag else '-'

    # Извлекаем Пакет / Package
    package_tag = product_page.select_one('#detail_infomation > table > tbody > tr:nth-child(6) > td:nth-child(4)')
    package = package_tag.get_text(strip=True) if package_tag else '-'

    # Извлекаем Характер продукции / product_description (HTML код)
    product_description_tag = product_page.select_one('#product_description')
    product_description = str(product_description_tag) if product_description_tag else '-'

    return [
        product_url, 'https://bwintools.com' + image1_url, 'https://bwintools.com' + image2_url,
        title, model_number, product_name, material, usage, hrc, application, ', '.join(high_lights), 
        color, workpiece, coating, feature, package, product_description
    ]

# Функция для парсинга товаров с одной страницы
def parse_products(page_url):
    products_page = get_soup(page_url)
    products = products_page.findAll('li', class_='item')
    data = []

    for product in products:
        url_tag = product.find('a', class_='image-all')
        product_url = url_tag['href'].strip() if url_tag else None

        if product_url:
            full_product_url = 'https://russian.bwintools.com' + product_url
            print(f'Парсим товар: {full_product_url}')
            data.append(parse_product_page(full_product_url))

    return data

# Открываем Excel файл для записи
with xlsxwriter.Workbook('Твердосплавные_вставки.xlsx') as workbook:
    
    # Функция для записи данных на отдельный лист
    def save_to_sheet(sheet_name, data):
        worksheet = workbook.add_worksheet(sheet_name)
        headers = ['URL товара', 'URL картинки 1', 'URL картинки 2', 'Название', 'Номер модели / Model Number',
                   'Название продукта / Product Name', 'Материал / Material', 'Использование / Usage',
                   'Твердость / HRC', 'Применение / Application', 'Высокий свет / High Light',
                   'Цвет / Color', 'Заготовка / Workpiece', 'Покрытие / coating', 'Особенность / Feature', 'Пакет / Package', 
                   'Характер продукции / product_description']
        worksheet.write_row(0, 0, headers)
        for row_num, row_data in enumerate(data, 1):
            worksheet.write_row(row_num, 0, row_data)

    # Парсим товары для каждого листа
    for page_num in range(1, 2):
        page_url = f'{main_url}supplier-3949298p{page_num}-u-drill-insert'
        print(f'Парсим страницу: {page_url}')
        drill_data = parse_products(page_url)
        save_to_sheet('U образная вставка для сверла', drill_data)

    for page_num in range(1, 3):
        page_url = f'{main_url}supplier-3949293p{page_num}-carbide-lathe-insert'
        print(f'Парсим страницу: {page_url}')
        lathe_data = parse_products(page_url)
        save_to_sheet('Вставка для токарного станка из карбида', lathe_data)  

    for page_num in range(1, 3):
        page_url = f'{main_url}supplier-3949278p{page_num}-cnc-carbide-inserts'
        print(f'Парсим страницу: {page_url}')
        cnc_carbide_inserts = parse_products(page_url)
        save_to_sheet('Вставки карбида cnc', cnc_carbide_inserts)

    for page_num in range(1, 3):
        page_url = f'{main_url}supplier-3949276p{page_num}-tungsten-carbide-inserts'
        print(f'Парсим страницу: {page_url}')
        tungsten_carbide_inserts = parse_products(page_url)
        save_to_sheet('Вставки карбида вольфрама', tungsten_carbide_inserts)  

    for page_num in range(1, 2):
        page_url = f'{main_url}supplier-3949294p{page_num}-carbide-grooving-insert'
        print(f'Парсим страницу: {page_url}')
        carbide_grooving_insert = parse_products(page_url)
        save_to_sheet('Карбид калибруя вставку', carbide_grooving_insert) 

    for page_num in range(1, 2):
        page_url = f'{main_url}supplier-3949296p{page_num}-carbide-threading-inserts'
        print(f'Парсим страницу: {page_url}')
        carbide_threading_inserts = parse_products(page_url)
        save_to_sheet('Карбид продевая нитку вставки', carbide_threading_inserts)

    for page_num in range(1, 3):
        page_url = f'{main_url}supplier-3949277p{page_num}-turning-carbide-inserts'
        print(f'Парсим страницу: {page_url}')
        turning_carbide_inserts = parse_products(page_url)
        save_to_sheet('Токарные твердосплавные пластины', turning_carbide_inserts)

    for page_num in range(1, 3):
        page_url = f'{main_url}supplier-3949285p{page_num}-milling-carbide-insert'
        print(f'Парсим страницу: {page_url}')
        milling_carbide_insert = parse_products(page_url)
        save_to_sheet('Фрезерная твердосплавная вставка', milling_carbide_insert)    

print("Парсинг завершен, данные сохранены в 'Твердосплавные_вставки.xlsx'.")