from http.server import HTTPServer, SimpleHTTPRequestHandler
import datetime
import pandas
import collections
import argparse
from jinja2 import Environment, FileSystemLoader, select_autoescape


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    parser = argparse.ArgumentParser()
    parser.add_argument('--path', type=str, default='wines.xlsx')
    args = parser.parse_args()
    path = args.path

    now = datetime.datetime.now()
    founded_date = 1920
    age = now.year - founded_date

    age_last = int(str(age)[-2:])
    age_suffix = "лет"
    if age_last == 1 or str(age)[-1] == '1':
        age_suffix = "год"
    if age_last in [2, 3, 4] or str(age)[-1] in ['2', '3', '4']:
        age_suffix = "года"
    if age_last > 10 and age_last <= 20:
        age_suffix = "лет"

    wines_df = pandas.read_excel(path,
                                 sheet_name='Лист1',
                                 keep_default_na="").to_dict(orient='records')

    wines = collections.defaultdict(list)
    for wine in wines_df:
        category = wine['Категория']
        wines[category].append(wine)

    rendered_page = template.render(
        age=age,
        age_suffix=age_suffix,
        wines=wines,
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
