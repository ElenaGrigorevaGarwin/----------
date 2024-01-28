import pandas as pd
import numpy as np

# Переменные
CNY = 15 # Конвертирование валюты
USD = 100  # Конвертирование валюты
months_of_stock = 5 # На сколько месяцев ТЗ
max_container_weight = 24570  # Вместимость контейнера по весу

# Загрузка данных
order = pd.read_excel('data/Заказ.xlsx')
code_infor = pd.read_excel('data/Код Инфор.xlsx')
quantity_per_box = pd.read_excel('data/шт в коробке.xlsx')
file_names = ['data/Продажи', 'data/Прогнозы', 'data/Остатки', 'data/Товары в пути']
sales, forecasts, balance,  invoice = [pd.read_excel(f'{file_name}.xlsx') for file_name in file_names]
retail_prices = pd.read_excel('data/Розничные цены.xlsx')

marged_data = pd.merge(order, code_infor[['Артикул', 'Код Инфор']].drop_duplicates(), on='Артикул', how='left')  # Объединение данных по столбцу "Артикул"
order['Код Инфор'] = marged_data['Код Инфор']

# Умножение "цена(Валюта)" на "Валюта" и сохранение в новую колонку "Цена(рубли)"
order['Цена(рубли)'] = order.apply(lambda row: row['цена(Валюта)'] * CNY if row['Валюта'] == 'CNY' else row['цена(Валюта)'] * USD if row['Валюта'] == 'USD' else row['цена(Валюта)'], axis=1)
order['Сумма(рубли)'] = order.apply(lambda row: row['Цена(рубли)'] * row['готовое к отгрузке кол-во (шт)'], axis=1)
marged_data2 = pd.merge(order, quantity_per_box[['Модель', 'штук в кор', 'нетто короба', 'брутто короба']], left_on='Артикул', right_on='Модель', how='left')
order[['штук в кор', 'нетто короба', 'брутто короба']] = marged_data2[['штук в кор', 'нетто короба', 'брутто короба']]

# Проверяем на соответствие месяцы в файле - Продажи
selected_sales = ['Октябрь 2022', 'Ноябрь 2022', 'Декабрь 2022', 'Январь 2023', 'Февраль 2023',
                  'Март 2023', 'Апрель 2023', 'Май 2023', 'Июнь 2023', 'Июль 2023', 'Август 2023', 'Сентябрь 2023'] # выбираем столбцы для суммирования в продажах

# Добавляем новый столбец 'Среднемесячные продажи'
sales['Среднемесячные продажи'] = round(sales[selected_sales].sum(axis=1) / 12)
marged_data3 = pd.merge(order, sales[['Код Инфор', 'Среднемесячные продажи']], on='Код Инфор', how='left')
order['Среднемесячные продажи'] = marged_data3['Среднемесячные продажи'].fillna(0).astype(int)
# Добавляем новый столбец 'Среднемесячные прогнозы'
forecasts['Среднемесячный прогноз'] = forecasts.apply(lambda row: row['Сумма прогноза за 6 мес'] / 6, axis=1)
forecasts['Среднемесячный прогноз'] = forecasts['Среднемесячный прогноз'].fillna(0).round().astype(int)
marged_data4 = pd.merge(order, forecasts[['Код_Инфор', 'Среднемесячный прогноз']], left_on='Код Инфор', right_on='Код_Инфор', how='left')
order['Среднемесячный прогноз'] = marged_data4['Среднемесячный прогноз'].fillna(0).astype(int)
marged_data5 = pd.merge(order, balance[['Товар.Код Инфор', 'Свободный остаток']], left_on='Код Инфор', right_on='Товар.Код Инфор', how='left')
order['Остаток'] = marged_data5['Свободный остаток'].fillna(0).astype(int)
marged_data6 = pd.merge(order,invoice[['Номенклатура.Код_Инфор', 'СУММА(КоличествоОстаток)']],  left_on='Код Инфор', right_on='Номенклатура.Код_Инфор', how='left')
order['Инвойс'] = marged_data6['СУММА(КоличествоОстаток)'].fillna(0).astype(int)
# проверка на наличие файла с сформированным контейнером от поставщика
try:
    formed_container = pd.read_excel('data/Сформированный контейнер.xlsx')
except FileNotFoundError:
    # Если файл не найден, создаем DataFrame с нужными столбцами и заполняем его нулями
    formed_container = pd.DataFrame({'Артикул Поставщика': order['Артикул Поставщика'], 'Готово к отгрузке': 0})

marged_data7 = pd.merge(order, formed_container[['Артикул Поставщика', 'Готово к отгрузке']], on='Артикул Поставщика', how='left')
order['Готово к отгрузке'] = marged_data7['Готово к отгрузке'].fillna(0).astype(int)
def custom_round(value):
  if np.isfinite(value):
    return int(np.floor(value))
  else:
    return 10
# Создаем пользовательскую функцию для округления вниз и преобразования в целое число с обработкой NaN и бесконечности
order['Месяцев ТЗ'] = (
    ((order['Остаток'] + order['Инвойс'] + order['Готово к отгрузке']) / order['Среднемесячные продажи'])
    .where((order['Среднемесячные продажи'] > 0) & ~ ((order['Среднемесячные продажи'] + order['Среднемесячный прогноз']) == 0),
           (order['Остаток'] + order['Инвойс'] + order['Готово к отгрузке']) / order['Среднемесячный прогноз'])
    .fillna(10)
    .apply(custom_round)
)

retail_prices['Розничная цена'] = retail_prices.apply(
    lambda row: 0 if pd.isna(row['USD']) and pd.isna(row['Рубли']) else
                  row['USD'] if pd.notna(row['USD']) else
                  row['Рубли'],
    axis=1
)
retail_prices['Проверка на валюту РЦ'] = retail_prices.apply(
    lambda row: 'USD' if pd.notna(row['USD']) else
                'Рубли' if pd.notna(row['Рубли']) else 'unknown',
    axis=1
)

marged_data8 = pd.merge(order, retail_prices[['Номенклатура.Код_Инфор', 'Розничная цена', 'Проверка на валюту РЦ']], left_on='Код Инфор', right_on='Номенклатура.Код_Инфор', how='left')
order[['Розничная цена', 'Проверка на валюту РЦ']] = marged_data8[['Розничная цена', 'Проверка на валюту РЦ']]
order['Розничная цена'] = order['Розничная цена'].fillna(0).astype(float)
order['Проверка на валюту РЦ'] = order['Проверка на валюту РЦ'].fillna('unknown')

order['ВП'] = (
    (
        ((order['Розничная цена'] * USD - order['Цена(рубли)']) * (order['готовое к отгрузке кол-во (шт)'] - order['Готово к отгрузке']))
        .where((order['Проверка на валюту РЦ'] == 'USD') & ~(order['Проверка на валюту РЦ'] == 'unknown'))
    )
    .fillna(
        (order['Розничная цена'] - order['Цена(рубли)']) * (order['готовое к отгрузке кол-во (шт)'] - order['Готово к отгрузке'])
    )
    .fillna(0)
)
# Добавляем условие перед операцией
condition = (order['готовое к отгрузке кол-во (шт)'] - order['Готово к отгрузке'] <= 0)
order = order.loc[~condition]  # Отфильтровываем строки, где условие истинно

order['Количество товара с запасом на 5 месяцев'] = order.apply(lambda row: row['Среднемесячные продажи'] * months_of_stock, axis=1)
order['Количество замороженного тавара'] = order.apply(lambda row: row['готовое к отгрузке кол-во (шт)'] - row['Готово к отгрузке'] - row['Количество товара с запасом на 5 месяцев'], axis=1)
order['Валова за единицу товара'] = order.apply(
    lambda row: 0 if row['Количество товара с запасом на 5 месяцев'] == 0 else
    row['ВП'] / (row['готовое к отгрузке кол-во (шт)'] - row['Готово к отгрузке']),
    axis=1
)
order['Замороженные деньги'] = order.apply(
    lambda row: 0 if row['Валова за единицу товара'] == 0 else
    row['Количество замороженного тавара'] * row['Валова за единицу товара'] * -1 if row['Количество замороженного тавара'] > 0 else
    0,
    axis=1
)
order['ВП товаров за минусом замороженных денег'] = order.apply(
    lambda row: 0 if row['Замороженные деньги'] + row['ВП'] < 0 else
    row['Замороженные деньги'] + row['ВП'],
    axis=1
)
order['Вес'] = order.apply(
    lambda row: np.ceil((row['готовое к отгрузке кол-во (шт)'] - row['Готово к отгрузке']) / row['штук в кор']) * row['брутто короба'] if pd.notna(row['штук в кор']) else
    (row['готовое к отгрузке кол-во (шт)'] - row['Готово к отгрузке']) * row['Вес по одной позиции'],
    axis=1
)
order['Объём'] = order.apply(
    lambda row: (row['готовое к отгрузке кол-во (шт)'] - row['Готово к отгрузке']) * row['Объём по одной позиции'],
    axis=1
)
order['коэффициент(ВП товаров за минусом замороженных денег/вес)'] = order.apply(
    lambda row: row['ВП товаров за минусом замороженных денег'] / row['Вес'],
    axis=1
)

# Сортировка товаров по приоритетам: месяцу обнуления и коэффициенту
order.sort_values(by=['Месяцев ТЗ', 'коэффициент(ВП товаров за минусом замороженных денег/вес)'], ascending=[True, False], inplace=True)
from typing_extensions import OrderedDict
def fill_containers(order, max_container_weight):
    containers = []  # Список для хранения контейнеров
    current_container = []  # Текущий контейнер
    current_weight = 0  # Текущий вес контейнера

    for index, row in order.iterrows():
        article = row['Артикул']
        weight = row['Вес']

        added_to_container = False

        # Проверяем каждый контейнер, начиная с первого
        for container in containers:
            total_weight = sum(w for _, w in container) + weight
            if total_weight <= max_container_weight:
                container.append((article, weight))
                added_to_container = True
                break

        # Если товар не поместился в существующие контейнеры, создаем новый
        if not added_to_container:
            containers.append([(article, weight)])

    return containers

result_containers = fill_containers(order, max_container_weight)

# Создаем DataFrame из списка контейнеров с добавлением столбца "Позиция"

df_combined = pd.concat([pd.DataFrame(container, columns=['Артикул', 'Вес']).assign(Контейнер=pos + 1) for pos, container in enumerate(result_containers)])
marged_data9 = pd.merge(order, df_combined, on='Артикул', how='left')
# Сохраняем в Excel
marged_data9.to_excel('Обновленный_Заказ.xlsx', index=False)