from bs4 import BeautifulSoup
import openpyxl
import string
import requests


class PageParser:
    def __init__(self, link):
        self.wb = openpyxl.load_workbook(filename="prices.xlsx")
        self.sheet = self.wb['Sheet1']
        self.head_url = link
        self.head_page = None
        self.current_category = None
        self.current_category_link = None
        self.current_page = None
        self.current_item = None
        self.cells_counter = 2
        self.string_of_chars = string.ascii_letters + "ЙЦУКЕНГШЩЗХЪФЫВПРОЛДЖЭЯЧСМИТЬБЮ"


    def get_catalog_name(self):
        tmp = PgParser.current_page.find("div", class_="body_text")
        return tmp.find("h1").text

    def parse_page(self):
        http_req = requests.get(self.head_url).text
        self.head_page = BeautifulSoup(http_req, 'lxml')

    def go_next_category(self):
        if self.current_category is None:
            self.current_category = self.head_page.find("div", class_="catalog-section")
        else:
            self.current_category = self.current_category.find_next_sibling()

        self.current_category_link = "http://ruletka.by" + self.current_category.a["href"]+"?view=list&limit=900"
        print(self.current_category_link)

    def go_to_category(self):
        page_req = requests.get(self.current_category_link).text
        self.current_page = BeautifulSoup(page_req, "lxml")

    def get_info(self):
        info_dict = {}

        if self.current_item is None:
            self.current_item = self.current_page.find("div", class_="catalog-item")
        else:
            self.current_item = self.current_item.find_next_sibling()

        title_link = self.current_item.find("div", class_="catalog-item-title")
        title = title_link.find("span").text
        info_dict["title"] = title

        price = self.current_item.find("span", class_="catalog-item-price")
        if price == None:
            info_dict["price"] = 0
        else:
            price = price.text
            price_val = ''
            for val in price:

                if val.isdigit() or val == '.':
                    price_val += val
            print(price_val[:-1])
            info_dict["price"] = float(price_val[:-1])

        article = self.current_item.find("div", class_="article").text
        final_article = ''

        for val in article:
            if val.isdigit() or val in '.-' or val in PgParser.string_of_chars:
                final_article += val


        info_dict["article"] = final_article

        print(info_dict)
        return info_dict

    def get_number_of_items(self):
        tmp = self.current_page.find("div", class_="count_items")
        return int(tmp.span.text)

    def move_data(self, info):

        cell_title = 'A'
        cell_article = 'B'
        cell_price = 'C'
        self.sheet[cell_title + str(self.cells_counter)] = info["title"]
        self.sheet[cell_article + str(self.cells_counter)] = info["article"]
        self.sheet[cell_price + str(self.cells_counter)] = info["price"]
        self.cells_counter += 1

    def move_category(self):
        self.sheet["A"+str(self.cells_counter)] = PgParser.get_catalog_name()
        self.cells_counter += 1


if __name__ == '__main__':
    number_of_categories = 8
    categories_counter = 0
    PgParser = PageParser("http://ruletka.by/catalog/")
    PgParser.parse_page()

    while categories_counter < number_of_categories:
        PgParser.current_page = None
        PgParser.current_item = None
        PgParser.go_next_category()
        PgParser.go_to_category()

        categories_counter += 1
        item_counter = 0

        PgParser.move_category()

        while item_counter < PgParser.get_number_of_items():
            PgParser.get_catalog_name()
            data_dict = PgParser.get_info()
            PgParser.move_data(data_dict)

            item_counter += 1
    PgParser.wb.save("prices.xlsx")






