import datetime
import pandas
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from collections import defaultdict


env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

start_day = datetime.datetime(year=1904, month=1, day=1, hour=0)
today = datetime.datetime.today()
delta_year = today.year - start_day.year

def get_year_form(delta_year):
    if delta_year % 100 in [11, 12, 13, 14]:
        return 'лет'
    last_digit = delta_year % 10
    if last_digit == 1:
        return 'год'
    elif last_digit in [2, 3, 4]:
        return 'года'
    else:
        return 'лет'

year_form = get_year_form(delta_year)

wine_table = pandas.read_excel('wine3.xlsx', sheet_name='Лист1')

wine_table = wine_table.rename(columns={
    'Категория': 'category',
    'Название': 'title',
    'Сорт': 'variety',
    'Цена': 'price',
    'Картинка': 'image',
    'Акция': 'promo'
})

wine_table['image'] = 'images/' + wine_table['image'].astype(str)
wine_table['variety'] = wine_table['variety'].where(wine_table['variety'].notna(), None)
wine_table['promo'] = wine_table['promo'].fillna('').str.strip()

wines_category = defaultdict(list)

for row in wine_table.itertuples(index=False):
    category = row.category.strip().lower()
    wine_info = {
        'title': row.title,
        'variety': row.variety,
        'price': row.price,
        'image': row.image,
        'promo': row.promo == 'Выгодное предложение'
    }
    wines_category[category].append(wine_info)

rendered_page = template.render(
    wines_category=wines_category,
    delta_year=delta_year,
    year_form=year_form
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
