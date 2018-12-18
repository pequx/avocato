# -*- coding: utf-8 -*-
"""
Module Docstring
"""

__author__ = "Maciej Piernikowski"
__version__ = "0.1.0"
__license__ = "MIT"

import argparse, csv, re, datetime, os, uuid, unicodedata
from openpyxl import load_workbook
from collections import OrderedDict
from fabulous.color import bold, highlight_red, highlight_green, green, italic, highlight_yellow

superatribute = 'superatribute'

class Category:
    target = 'category.csv'
    categories = {
        'level_1': {},
        'level_2': {}
    }
    mappings_processed_count = 0
    categories_processed_count = {
        'level_1': 0,
        'level_2': 0
    }

    def __init__(self):
        Logger.highlight('Processing of categories...')
        self.categories['level_1']['root'] = {
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
        self.categories['level_1']['product-bundles'] = {
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
        Logger.msg('Processed two system categories: '+bold('root,product-bundle'))

    def process(self):
        for mapping in Processor.categories:
            level_1 = str(lowerCase(mapping['level_1']))
            if not level_1 in self.categories['level_1']:
                self.categories['level_1'][level_1] = {
                    'category_key': lowerCase(mapping['level_1']),
                    'parent_category_key': 'demoshop',
                    'name.de_DE': mapping['level_1'],
                    'name.en_US': mapping['level_1'],
                    'meta_title.de_DE': mapping['level_1'],
                    'meta_title.en_US': mapping['level_1'],
                    'meta_description.de_DE': mapping['level_1'],
                    'meta_description.en_US': mapping['level_1'],
                    'meta_keywords.de_DE': mapping['level_1'],
                    'meta_keywords.en_US': mapping['level_1'],
                    'is_active': 1,
                    'is_in_menu': 1,
                    'is_clickable': 1,
                    'is_searchable': 1,
                    'is_root': 0,
                    'is_main': 1,
                    'node_order': 40,
                    'template_name': 'Catalog + CMS Block'   
                }
                self.categories_processed_count['level_1'] += 1
                Logger.msg('Processing of the category level 1: '+bold(level_1)+' completed with parent: '+bold('demoshop'))
            level_2 = str(lowerCase(mapping['level_2']))
            if level_2 in self.categories['level_2']: level_2 = level_2+'-'+str(uuid.uuid4())
            self.categories['level_2'][level_2] = {
                'category_key': lowerCase(mapping['level_2']),
                'parent_category_key': level_1,
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
            self.categories_processed_count['level_2'] += 1
            Logger.msg('Processing of the category level 2: '+bold(level_2)+' completed with parent: '+bold(level_1))
            self.mappings_processed_count += 1
        Logger.highlight('Processing of categories completed.')
        Logger.output('caegories', self.target)
        Logger.msg('Processed '+bold(str(self.mappings_processed_count))+
                ' category mappings providing '+bold(str(self.categories_processed_count['level_1']))+
                ' level 1 categories, and '+bold(str(self.categories_processed_count['level_2']))+
                ' level 2 categories.')
        # field_names = self.categories['product-bundles'].keys()
        # with open(self.target, 'wb') as output_file:
        #     dict_writer = csv.DictWriter(output_file, field_names, delimiter=',')
        #     dict_writer.writeheader()
        #     # dict_writer.writerows(self.categories)
        #     for category in self.categories:
        #         dict_writer.writerow(self.categories[category])
        # return args
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
                self.last_abstract = str(current['Product SKU'])
                Logger.update(current['Product SKU'])
        Logger.highlight('Processing of abstract products completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_abstracts[self.last_abstract].keys())
        writer.process()
        writer.write()
        Logger.output('abstract products', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))
class ProductConcrete:
    target = 'product_concrete.csv'
    product_concretes = {}
    product_concretes_orphaned = {}
    product_processed_count = 0

    def process(self):
        Logger.highlight('Processing of concrete products...')
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
                    self.product_processed_count += 1
                    msg = 'paired concrete SKU with parent abstract SKU '+ bold(parent['abstract_sku'])
                    Logger.update(current['Product SKU'], msg)
                except KeyError:
                    self.product_concretes_orphaned[str(current['Product SKU'])] = current
        Logger.highlight('Processing of concrete products completed.')
        Logger.output('abstract products', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))
class ProductAbstractStore:
    target = 'product_abstract_store.csv'
    product_abstract_stores = []
    product_processed_count = 0
    store_processed_count = 0
    stores_avaiable = ['DE', 'AT', 'US']

    def process(self):
        Logger.highlight('Processing of abstract product stores...')
        for product in ProductAbstract.product_abstracts:
            for store in self.stores_avaiable:
                self.product_abstract_stores.append({ 
                    'product_abstract_sku': product, 
                    'store_name': store 
                })
                self.store_processed_count += 1
                msg = 'paired abstract SKU with parent store name '+ bold(store)
                Logger.update(product, msg)
            self.product_processed_count += 1
        del product
        Logger.highlight('Processing of abstract product stores completed.')
        Logger.output('abstract product stores', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))
class ProductAttributeKey:
    target = 'product_attribute_key.csv'
    product_attribute_keys = {}
    keys_processed_count = 0

    def process(self):
        Logger.highlight('Processing of product attribute keys...')
        self.product_attribute_keys[superatribute] = { 'is_super': True }
        self.keys_processed_count += 1
        Logger.msg('Processed attribute key '+ bold(superatribute))
        Logger.highlight('Processing of product attribute keys completed.')
        Logger.output('product attribute keys', self.target)
        Logger.msg('Processed '+bold(str(self.keys_processed_count))+' product attribute keys')
class ProductImage:
    target = 'product_image.csv'
    product_images = []
    locales_avaiable = ['DE', 'US']
    products_processed_count = 0
    images_processed_count = 0
    stores_processed_count = 0
    missed_products = []
    products = {}
    current = {}

    def process_store(self, product):
        for store in self.locales_avaiable:
            image = Processor.products[product]['Featured Image']
            if image != '':
                self.images_processed_count += 1
                if 'abstract_sku' in self.current: 
                    self.product_images.append({
                        'abstract_sku': self.current['abstract_sku'],
                        'concrete_sku': '',
                        'image_set_name': 'default',
                        'external_url_large': '',
                        'external_url_small': Processor.products[product]['Featured Image'],
                        'locale': getLocale(store)
                    })
                    msg = 'paired abstract SKU with children image '+bold(Processor.products[product]['Featured Image'])+' and store '+bold(store)
                    Logger.update(product, msg)
                    self.products_processed_count += 1
                elif 'concrete_sku' in self.current:
                    self.product_images.append({
                        'abstract_sku': '',
                        'concrete_sku': self.current['concrete_sku'],
                        'image_set_name': 'default',
                        'external_url_large': Processor.products[product]['Featured Image'],
                        'external_url_small': Processor.products[product]['Featured Image'],
                        'locale': getLocale(store)
                    })
                    self.products_processed_count += 1
                    msg = 'paired concrete SKU with children image '+bold(Processor.products[product]['Featured Image'])+' and store '+bold(store)
                    Logger.update(product, msg)
            elif image == '':
                self.missed_products.append(product)
                msg = 'Processing of the SKU: '+product+' failed, with message '+italic('no product image defined')
                Logger.warning(msg)
            self.stores_processed_count += 1

    def process(self):
        Logger.highlight('Processing of product images...')
        for product in ProductAbstract.product_abstracts:
            self.current = ProductAbstract.product_abstracts[product]
            self.process_store(product)
        for product in ProductConcrete.product_concretes:
            self.current = ProductConcrete.product_concretes[product]
            self.process_store(product)
        Logger.highlight('Processing of product images keys completed.')
        Logger.output('product images', self.target)
        Logger.msg('Processed '+bold(str(self.products_processed_count))+' SKUs with '+bold(str(self.images_processed_count))+' images in '+ bold(str(len(self.locales_avaiable)))+' stores from '+bold(str(len(Processor.products)))+' imported products.')
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
       Logger.highlight('Loading of product labels completed.')
       Logger.output('product labels', self.target)
       Logger.msg('Processed '+bold(str(len(self.product_labels)))+' product labels.')
class ProductManagmentAttribute:
    target = 'product_management_attribute.csv'
    product_management_attributes = []
    products_processed_count = 0
    attributes_processed_count = 0

    def getAttributes(self):
        attributes = []
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            attribute = current['value_1']
            self.products_processed_count += 1
            if attribute not in attributes:
                attributes.append(attribute)
                self.attributes_processed_count += 1
                Logger.msg('Processing of the SKU: '+bold(product)+' completed with message: '+italic('collected a new attribute: '+bold(attribute)))
        return ','.join(list(OrderedDict.fromkeys(attributes)))

    def process(self):
        Logger.highlight('Processing of product managment attributes...')
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
        Logger.highlight('Processing of product managment attributes completed.')
        Logger.output('product managment attributes', self.target)
        Logger.msg('Processed '+bold(superatribute)+' with values '+bold(attributes)+' from '+bold(self.products_processed_count)+' concrete products.')
class ProductPrice:
    target = 'product_price.csv' 
    product_prices = []
    missed_product_prices = []
    products_processed_count = 0
    prices_processed_count = 0

    def process(self):
        Logger.highlight('Processing of product prices...')
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            try: 
                price = float(Processor.products[product]['Price'])
                tax = 0.2 * price
                net =  int((price - tax) * 100)
                gross = int(price * 100)
                store = 'DE'
                self.product_prices.append({
                    'abstract_sku': '',
                    'concrete_sku': current['concrete_sku'],
                    'price_type': 'DEFAULT',
                    'store': store,
                    'currency': 'EUR',
                    'value_net': net,
                    'value_gross': gross,
                    'price_data.volume_prices': ''
                })
                self.products_processed_count += 1
                if gross < net:
                    self.missed_product_prices.append(product)
                    Logger.warning('Processing of the SKU: '+product+' failed, with message: '+italic('net price > gross price.'))
                    continue
                self.prices_processed_count += 2
                Logger.msg('Processing of the SKU: '+bold(product)+' completed with the message: '+italic('paried concrete product with net/gross prices: '+bold([net, gross])+' for store: '+bold(store)))
            except KeyError: 
                self.missed_product_prices.append(product)
                Logger.warning('Processing of the SKU: '+product+' failed, with message: '+italic('no product price defined.'))
        Logger.highlight('Processing of product prices completed.')
        Logger.output('product prices', self.target)
        Logger.msg('Processed ' +bold(self.products_processed_count)+' SKUs with '+bold(self.prices_processed_count)+' prices (net+gross) for store: '+bold(store))
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
    categories = []
    
    def __init__(self, args):
        """ Main entry point of the app """
        print "hello world", args
        workbook = load_workbook(filename = args.filename, read_only=args.read_only)
        self.data_product_export = self.process_workbook(workbook['Product Export'])
        self.data_product_meta = self.process_workbook(workbook['Product Meta Data'])
        del workbook
    def hydrate(self):
        missed_ids = [] # products without id
        for row in self.data_product_export['rows']:
            if row > 1: 
                current_row = self.data_product_export['rows'][row]
                product = {}
                for index, value in enumerate(current_row):
                    product[self.data_product_export['headers'][index]] = value
                sku = product['Product SKU']
                if sku != '': self.products[str(product['Product SKU'])] = product
                if sku == '': missed_ids.append(product['Product ID'])
                del current_row
        del product, index, value, row
        for product in self.products:
            current = self.products[product]
            split = current['Category'].split('>')
            # if split[0] != 'Uncategorized': category_mappings[product] = { split[0]: split[1] }
            if split[0] != 'Uncategorized': self.categories.append({ 'level_1': split[0], 'level_2': split[1] })
            del current
        del product, split
        category_count = {}
        category_index = 0
        while category_index < len(self.categories):
            if self.categories[category_index]['level_2'] in category_count:
                del self.categories[category_index]
            else:
                category_count[self.categories[category_index]['level_2']] = 1
                category_index += 1
        del category_count, category_index
        missed_rows = {} # rows without data
        missed_matches = [] # meta prodocuts not mached to export products
        for row in self.data_product_meta['rows']:
            if row > 1:
                current_row = self.data_product_meta['rows'][row]
                meta_product = {}
                for index, value in enumerate(current_row):
                    meta_product[self.data_product_meta['headers'][index]] = value
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
        del current_row, current_product, row, sku
    def run(self):
        Category().process()
        ProductAbstract().process()
        ProductConcrete().process()
        ProductAbstractStore().process()
        ProductAttributeKey().process()
        ProductImage().process()
        ProductLabel().process()
        ProductManagmentAttribute().process()
        ProductPrice().process()
        Logger.highlight('Processing complete.')
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
    def update(sku, msg=False):
        if msg is False: print(green('Processing of the SKU: ' + bold(str(sku)) + ' completed.'))
        else: print(green('Processing of the SKU: ' + bold(str(sku)) + ' completed with message: ' + italic(msg)))
    @staticmethod
    def summary(num, total, msg=False):
        if msg is False: print(green('Processed ' + bold(str(num)) + ' SKUs of ' + bold(str(total)) + ' imported products.'))
        else: print(green('Processed ' + bold(str(num)) + ' SKUs of ' + bold(str(total)) + ' imported products with message: ') + italic(msg))
    @staticmethod
    def output(type, target):
        path = os.path.dirname(os.path.realpath(__file__))
        print(green('Generated the CSV file with '+bold(type)+' located in '+bold(path+'/'+target)))
    @staticmethod
    def msg(msg):
        print(green(msg))
    @staticmethod
    def warning(msg):
        print(highlight_yellow(msg))
class Writer: 
    def __init__(self, target):
        self.file = target
        self.path = os.path.dirname(os.path.realpath(__file__))
        self.target = target
    def get_fieldnames(self, headers):
        with open(self.path+'/'+self.file, mode='w') as csv_file:
            self.writer = csv.DictWriter(csv_file, fieldnames=headers)
            if len(headers) > 0: self.writer.writeheader()
    def process(self):
        if self.target == 'product_abstract.csv':
            self.queue = ProductAbstract.product_abstracts
            ProductAbstract.product_abstracts = True
        for item in self.queue:
            current_item = self.queue[item]
            for key in current_item:
                value = str(current_item[key])
                value.replace('', '')
                value = unicodedata.normalize('NFKD', value)
                current_item[key] = value
                print value;
            print 'ja'
    def write(self):
        if self.target == 'product_abstract.csv':
            for product in self.queue:
                current = self.queue[product]
                self.writer.writerow(current)
                del current

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
    processor = Processor(args)
    processor.hydrate()
    processor.run()

