# -*- coding: utf-8 -*-
"""
Module Docstring
"""

__author__ = "Maciej Piernikowski"
__version__ = "0.1.0"
__license__ = "MIT"

import argparse, csv, re, datetime
from openpyxl import load_workbook
from collections import OrderedDict
from fabulous.color import bold, highlight_red, highlight_green, green

superatribute = 'superatribute'

class Category:
    target = 'category.csv'
    categories = {}

    def __init__(self):
        self.categories['root'] = {
            'category_key': 'demoshop',
            'name.de_DE': 'Demoshop',
            'name.en_US': 'Demoshop',
            'meta_title.de_DE': 'Demoshop',
            'meta_title.en_US': 'Demoshop',
            'meta_description.de_DE': 'Deutsche Version des Demoshop',
            'meta_description.en_US': 'English version of Demoshop',
            'meta_keywords.de_DE': 'Deutsche Version des Demoshop',
            'meta_keywords.en_US': 'English version of Demoshop',
            'is_active': 1,
            'is_in_menu': 1,
            'is_clickable': 1,
            'is_searchable': 0,
            'is_root': 1,
            'is_main': 1,
            'template_name': 'Catalog (default)'
        }
        self.categories['product-bundles'] = {
            'category_key': 'product-bundles',
            'parent_category_key': 'demoshop',
            'name.de_DE': 'Product Bundles',
            'name.en_US': 'Product Bundles',
            'meta_title.de_DE': 'Product Bundles',
            'meta_title.en_US': 'Product Bundles',
            'meta_description.de_DE': 'Diese Produkte ha√üben mehr als eine Variante.',
            'meta_description.en_US': 'These are multiple products bundled to a new product.',
            'meta_keywords.de_DE': 'Product Bundles',
            'meta_keywords.en_US': 'Product Bundles',
            'is_active': 1,
            'is_in_menu': 0,
            'is_clickable': 1,
            'is_searchable': 0,
            'is_root': 1,
            'is_main': 1,
            'node_order': 50,
            'template_name': 'Catalog (default)'   
        }

    def process(self, args):
        print 'Processing: category mapping...'
        for mapping in args:
            self.categories[str(lowerCase(mapping['level_2']))] = {
                'category_key': lowerCase(mapping['level_1']),
                'parent_category_key': 'demoshop',
                'name.de_DE': mapping['level_2'],
                'name.en_US': mapping['level_2'],
                'meta_title.de_DE': mapping['level_2'],
                'meta_title.en_US': mapping['level_2'],
                'meta_description.de_DE': mapping['level_2'],
                'meta_description.en_US': mapping['level_2'],
                'meta_keywords.de_DE': mapping['level_2'],
                'meta_keywords.en_US': mapping['level_2'],
                'is_active': 1,
                'is_in_menu': 1,
                'is_clickable': 1,
                'is_searchable': 1,
                'is_root': 0,
                'is_main': 1,
                'node_order': 40,
                'template_name': 'Catalog + CMS Block'   
            }
            self.categories[str(lowerCase(mapping['level_2']))] = {
                'category_key': lowerCase(mapping['level_2']),
                'parent_category_key': camelCase(mapping['level_1']),
                'name.de_DE': mapping['level_2'],
                'name.en_US': mapping['level_2'],
                'meta_title.de_DE': mapping['level_2'],
                'meta_title.en_US': mapping['level_2'],
                'meta_description.de_DE': mapping['level_2'],
                'meta_description.en_US': mapping['level_2'],
                'meta_keywords.de_DE': mapping['level_2'],
                'meta_keywords.en_US': mapping['level_2'],
                'is_active': 1,
                'is_in_menu': 1,
                'is_clickable': 1,
                'is_searchable': 1,
                'is_root': 0,
                'is_main': 1,
                'node_order': 30,
                'template_name': 'Catalog + CMS Block'
            }
        field_names = self.categories['product-bundles'].keys()
        with open(self.target, 'wb') as output_file:
            dict_writer = csv.DictWriter(output_file, field_names, delimiter=',')
            dict_writer.writeheader()
            # dict_writer.writerows(self.categories)
            for category in self.categories:
                dict_writer.writerow(self.categories[category])
        return args
class ProductAbstract:
    target = 'product_abstract.csv'
    product_abstracts = {}
    product_new_threshold_days = 31 # number of days after which product will not be marked as `new`
    product_processed_count = 0

    @staticmethod
    def category_key(args):
        return str(lowerCase(args[-1]))

    @staticmethod
    def is_featured(args):
        if args == 'Yes': return 1
        return 0
    
    @staticmethod
    def url(iso, current):
        match = re.compile('/product/.*').findall(current['Product URI'])
        if iso == 'de_DE': return match[0]
        if iso == 'en_US': return '/en' + match[0] 

    def process(self):
        Logger.highlight('Processing of abstract products...')
        for product in Processor.products:
            current = Processor.products[product]
            if current['Parent ID'] == '': # only true abstracts
                self.product_abstracts[str(current['Product SKU'])] = {
                    'category_key': self.category_key(current['Category'].split('>')),
                    'category_product_order': 2,
                    'abstract_sku': current['Product SKU'],
                    'tax_set_name': 'Standard Taxes',
                    'name.de_DE': current['Product Name'],
                    'name.en_US': current['Product Name'],
                    'description.de_DE': current['Description'],
                    'description.en_US': current['Description'],
                    'url.de_DE': self.url('de_DE', current),
                    'url.en_US': self.url('en_US', current),
                    'meta_title.de_DE': current['Post Title'],
                    'meta_title.en_US': current['Post Title'],
                    'meta_keywords.de_DE': current['Slug'].replace('-',' '),
                    'meta_keywords.en_US': current['Slug'].replace('-',' '),
                    'meta_description.de_DE': current['Description'],
                    'meta_description.en_US': current['Description'],
                    'is_featured': self.is_featured(current['Featured']),
                    # 'attribute_key_1': 'variant',
                    # 'value_1': current['Ebay ean'],
                    # 'attribute_key_1.de_DE': '',
                    # 'value_1.de_DE': '',
                    # 'attribute_key_1.en_US': '',
                    # 'value_1.en_US': '',
                    'color_code': '#FFFFFF',
                    'new_from': current['Product Published'].strftime('%Y-%m-%d %H:%M:%S.%f'), # 2018-08-01 00:00:00.000000
                    'new_to': (current['Product Published'] +
                         datetime.timedelta(days=self.product_new_threshold_days)).strftime('%Y-%m-%d %H:%M:%S.%f')
                }
                self.product_processed_count += 1
                Logger.update(current['Product SKU'])
        Logger.highlight('Processing of abstract products completed.')
        Logger.summary(self.product_processed_count, len(Processor.products))
class ProductConcrete:
    target = 'product_concrete.csv'
    product_concretes = {}
    product_concretes_orphaned = {}

    def process(self):
        for product in Processor.products:
            current = Processor.products[product]
            if current['Parent SKU'] != '': # only true concretes
                try:
                    parent = ProductAbstract.product_abstracts[str(current['Parent SKU'])]
                    self.product_concretes[str(current['Product SKU'])] = {
                        'abstract_sku': parent['abstract_sku'],
                        'old_sku': '',
                        'concrete_sku': current['Product SKU'],
                        'name.de_DE': current['Product Name'],
                        'name.en_US': current['Product Name'],
                        'description.de_DE': parent['description.de_DE'],
                        'description.en_US': parent['description.en_US'],
                        'is_searchable.de_DE': True,
                        'is_searchable.en_US': True,
                        'bundled': '',
                        'is_quantity_splittable': False,
                        'attribute_key_1': superatribute,
                        'attribute_key_1.de_DE': 'Jewellery',
                        'attribute_key_1.en_US': 'Jewellery',
                        'value_1': current['Attribute pa jewellery'],
                        'value_1.de_DE': current['Attribute pa jewellery'],
                        'value_1.en_US': current['Attribute pa jewellery']                 
                    }
                except KeyError:
                    self.product_concretes_orphaned[str(current['Product SKU'])] = current
        print 'ja'
class ProductAbstractStore:
    target = 'product_abstract_store.csv'
    product_abstract_stores = []
    stores_avaiable = ['DE', 'AT', 'US']

    def process(self):
        for product in ProductAbstract.product_abstracts:
            for store in self.stores_avaiable:
                self.product_abstract_stores.append({ 'product_abstract_sku': product, 'store_name': store })
        del product
        print 'ja'
class ProductAttributeKey:
    target = 'product_attribute_key.csv'
    product_attribute_keys = {}

    def process(self):
        self.product_attribute_keys[superatribute] = { 'is_super': True }
        print 'ja'
class ProductImage:
    target = 'product_image.csv'
    product_images = []
    locales_avaiable = ['DE', 'US']
    products = {}
    current = {}

    def __init__(self):
        self.products = products

    def process_store(self):
        for store in self.locales_avaiable:
            self.product_images.append({
                'abstract_sku': self.current['abstract_sku'],
                'concrete_sku': '',
                'image_set_name': 'default',
                'external_url_large': self.products[product]['Featured Image'],
                'external_url_small': self.products[product]['Featured Image'],
                'locale': getLocale(store)
            })

    def process(self):
        for product in ProductAbstract.product_abstracts:
            self.current = ProductAbstract.product_abstracts[product]
            self.process_store(product)
        for product in ProductConcrete.product_concretes:
            self.current = ProductConcrete.product_concretes[product]
            self.process_store(product)
        print 'ja'
class ProductLabel:
    target = 'product_label.csv'
    product_labels = {
        'TOP': {
            'name': 'TOP',
            'is_active': True,
            'is_dynamic': False,
            'is_exclusive': False,
            'front_end_reference': 'top',
            'valid_from': '',
            'valid_to': '',
            'name.de_DE': 'Top',
            'name.en_US': 'Top',
            'product_abstract_skus': ''
        },
        'NEW': {
            'name': 'NEW',
            'is_active': True,
            'is_dynamic': True,
            'is_exclusive': False,
            'front_end_reference': 'new',
            'valid_from': '',
            'valid_to': '',
            'name.de_DE': 'Neu',
            'name.en_US': 'New',
            'product_abstract_skus': ''
        },
        'SALE': {
            'name': 'NEW',
            'is_active': True,
            'is_dynamic': True,
            'is_exclusive': False,
            'front_end_reference': 'sale',
            'valid_from': '',
            'valid_to': '',
            'name.de_DE': 'SALE %',
            'name.en_US': 'SALE %',
            'product_abstract_skus': ''
        }
    }

    def process(self):
        print 'Processing: product labels...'
class ProductManagmentAttribute:
    target = 'product_management_attribute.csv'
    product_management_attributes = []

    @staticmethod
    def getAttributes():
        attributes = []
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            attributes.append(current['value_1'])
        return ','.join(list(OrderedDict.fromkeys(attributes)))

    def process(self):
        attributes = self.getAttributes()
        self.product_management_attributes.append({
            'key': superatribute,
            'input_type': 'text',
            'allow_input': 'yes',
            'is_multiple': 'yes',
            'values': attributes,
            'key_translation.en_US': attributes,
            'key_translation.de_DE': attributes,
            'value_translations.en_US': attributes,
            'value_translations.de_DE': attributes
        })
        print 'ja'
class ProductPrice:
    target = 'product_price.csv' 
    product_prices = []
    missed_product_prices = []
    products = {}

    def __init__(self, products):
        self.products = products

    def process(self):
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            try: 
                price = float(self.products[product]['Price'])
                tax = 0.2 * price
                self.product_prices.append({
                    'abstract_sku': '',
                    'concrete_sku': current['concrete_sku'],
                    'price_type': 'DEFAULT',
                    'store': 'DE',
                    'currency': 'EUR',
                    'value_net': int((price - tax) * 100),
                    'value_gross': int(price * 100),
                    'price_data.volume_prices': ''
                })
            except KeyError: 
                self.missed_product_prices.append(product)
        print 'ja'   
class ProductStock:
    target = 'product_stock.csv'
    product_stocks = {}

    def process(self):
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            self.product_stocks[product] = {
                'concrete_sku': '',
                'name': 'Warehouse1',
                'quantity': '',
                'is_never_out_of_stock': False,
                'is_bundle': False
            }

class Processor:
    products = {}
    
    def __init__(self, args):
        """ Main entry point of the app """
        print "hello world", args
        workbook = load_workbook(filename = args.filename, read_only=args.read_only)
        data_product_export = self.process_workbook(workbook['Product Export'])
        data_product_meta = self.process_workbook(workbook['Product Meta Data'])
        del workbook
        missed_ids = [] # products without id
        for row in data_product_export['rows']:
            if row > 1: 
                current_row = data_product_export['rows'][row]
                product = {}
                for index, value in enumerate(current_row):
                    product[data_product_export['headers'][index]] = value
                sku = product['Product SKU']
                if sku != '': self.products[str(product['Product SKU'])] = product
                if sku == '': missed_ids.append(product['Product ID'])
                del current_row
        del product, index, value, data_product_export, row
        category_mappings = []
        for product in self.products:
            current = self.products[product]
            split = current['Category'].split('>')
            # if split[0] != 'Uncategorized': category_mappings[product] = { split[0]: split[1] }
            if split[0] != 'Uncategorized': category_mappings.append({ 'level_1': split[0], 'level_2': split[1] })
            del current
        del product, split
        category_count = {}
        category_index = 0
        while category_index < len(category_mappings):
            if category_mappings[category_index]['level_2'] in category_count:
                del category_mappings[category_index]
            else:
                category_count[category_mappings[category_index]['level_2']] = 1
                category_index += 1
        del category_count, category_index
        Category().process(category_mappings)
        del category_mappings
        missed_rows = {} # rows without data
        missed_matches = [] # meta prodocuts not mached to export products
        for row in data_product_meta['rows']:
            if row > 1:
                current_row = data_product_meta['rows'][row]
                meta_product = {}
                for index, value in enumerate(current_row):
                    meta_product[data_product_meta['headers'][index]] = value
                del index, value
                sku = meta_product['Product SKU']
                if sku != '': 
                    try: 
                        current_product = self.products[str(sku)]
                        self.products[str(sku)].update(meta_product)
                    except KeyError: missed_matches.append(current_product['Product ID'])
                    del meta_product
                if sku == '': 
                    missed_rows[row] = current_row
                    continue
        del current_row, current_product, data_product_meta, row, sku
        ProductAbstract().process()
        ProductConcrete().process()
        ProductAbstractStore().process()
        ProductAttributeKey().process()
        ProductImage().process()
        ProductLabel().process()
        ProductManagmentAttribute().process()
        ProductPrice().process()
        # for product in ProductAbstract.product_abstracts:
        #     del products[product]
        # for product in ProductConcrete.product_concretes:
        #     del products[product]
        # del product
        # if products == ProductConcrete.product_concretes_orphaned:
        #     missed_products = products
        #     print 'bleh'
        #     del ProductConcrete.product_concretes_orphaned
        # del products
        # print 'missed_products'

    @staticmethod
    def process_workbook(sheet):
        row_count = 0
        cell_count = 0
        values = {}
        for row in sheet.rows:
            cols = []
            for cell in row:
                if cell.value is not None: cols.append(cell.value)
                elif cell.value is None: cols.append('')
                cell_count += 1
                del cell
            row_count += 1
            values[row_count] = cols
            del cols, row
        headers = values[1]
        del values[1] # removes first row as it contains headers
        return { 'headers': headers, 'rows': values }
class Logger:
    @staticmethod
    def highlight(msg):
        print(highlight_green(msg))
    @staticmethod
    def update(sku):
        print(green('Processing of the SKU: ' + bold(str(sku)) + ' completed.'))
    @staticmethod
    def summary(num, total, msg=False, ):
        if msg is False: print(green('Processed ' + bold(str(num)) + ' SKUs of ' + bold(str(total)) + ' imported products.'))
        if msg is True: print(green('Processed ' + bold(str(num)) + ' SKUs of ' + bold(str(total)) + ' imported products with message: ') +  bold(message))

def upperCase(string):
    output = string.replace('-',' ').upper()
    return output
def camelCase(string, space=False):
    output = ''.join(x for x in string.title() if x.isalnum())
    if space: return output[0].lower() + ' ' + output[1:]
    return output[0].lower() + output[1:]
def lowerCase(string):
    output = string.replace(' ', '-').lower()
    return output
def getLocale(store):
    if store == 'DE': return 'de_DE'
    if store == 'US': return 'en_EN'
    if store == 'AT': return 'de_DE'

if __name__ == "__main__":
    """ This is executed when run from the command line """
    parser = argparse.ArgumentParser()

    # Required positional argument
    parser.add_argument("filename", help="Required positional argument")

    # Optional argument flag which defaults to False
    parser.add_argument("-r", "--read-only", action="store", default=True)

    # Optional argument which requires a parameter (eg. -d test)
    parser.add_argument("-n", "--name", action="store", dest="name")

    # Optional verbosity counter (eg. -v, -vv, -vvv, etc.)
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Verbosity (-v, -vv, etc)")

    # Specify output of "--version"
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s (version {version})".format(version=__version__))

    args = parser.parse_args()
    Processor(args)

