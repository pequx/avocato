# -*- coding: utf-8 -*-
"""
Module Docstring
"""

__author__ = "Maciej Piernikowski <maciej@piernikowski.net>"
__version__ = "0.3.0"
__license__ = "MIT"

import argparse, csv, re, datetime, os, uuid, unicodedata, logging
from openpyxl import load_workbook
from collections import OrderedDict
from fabulous.color import bold, highlight_red, highlight_green, green, italic, highlight_yellow

superatribute = 'superatribute'


class CategoryTemplate:
    target = 'category_template.csv'
    category_templates = {}
    templates_processed_count = 0

    def __init__(self):
        Logger.highlight('Processing of category templates...')
        self.category_templates['Catalog (default)'] = {
            'template_name': 'Catalog (default)',
            'template_path': '@CatalogPage/views/catalog/catalog.twig'
        }
        self.category_templates['Catalog + CMS Block'] = {
            'template_name': 'Catalog + CMS Block',
            'template_path': '@CatalogPage/views/catalog-with-cms-block/catalog-with-cms-block.twig'
        }
        self.category_templates['CMS Block'] = {
            'template_name': 'CMS Block',
            'template_path': '@CatalogPage/views/simple-cms-block/simple-cms-block.twig'
        }
        self.templates_processed_count += 3

    def process(self):
        Logger.highlight('Processing of category templates completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.category_templates['Catalog (default)'].keys())
        writer.process()
        writer.write()
        Logger.output('category templates', self.target)


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
            'is_searchable': 1,
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
            'is_root': 0,
            'is_main': 1,
            'node_order': 50,
            'template_name': 'Catalog (default)'   
        }
        Logger.msg('Processed two system categories: '+bold('root,product-bundle'))

    def process(self):
        for mapping in Processor.categories:
            level_1 = str(lowerCase(mapping['level_1']))
            if level_1 not in self.categories['level_1']:
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
                self.last_category = level_1
                Logger.msg('Processing of the category level 1: '+bold(level_1)+' completed with parent: '+bold('demoshop'))
            level_2 = str(lowerCase(mapping['level_2']))
            if level_2 in self.categories['level_2']:
                level_2 = level_2+'-'+str(uuid.uuid4())
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
        writer = Writer(self.target)
        writer.get_fieldnames(self.categories['level_1'][self.last_category].keys())
        writer.process()
        writer.write()
        Logger.output('categories', self.target)
        Logger.msg('Processed '+bold(str(self.mappings_processed_count))
                   + ' category mappings providing '+bold(str(self.categories_processed_count['level_1']))
                   + ' level 1 categories, and '+bold(str(self.categories_processed_count['level_2']))
                   + ' level 2 categories.')


class ProductAbstract:
    target = 'product_abstract.csv'
    product_new_threshold_days = 31  # number of days after which product will not be marked as `new`
    product_abstracts = {}

    def __init__(self):
        Logger.highlight('Processing of abstract products...')
        self.product_processed_count = 0
        self.last_abstract = None

    @staticmethod
    def category_key(string):
        return str(lowerCase(string[-1]))

    @staticmethod
    def is_featured(val):
        if val == 'Yes':
            return 1
        return 0
    
    @staticmethod
    def url(iso, current):
        match = re.compile('/product/.*').findall(current['Product URI'])
        if iso == 'de_DE':
            return match[0]
        if iso == 'en_US':
            return '/en' + match[0]

    def process(self):
        for product in Processor.products:
            current = Processor.products[product]
            if current['Parent ID'] == '':  # only true abstracts
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
                    'new_from': current['Product Published'].strftime('%Y-%m-%d %H:%M:%S.%f'),  # 2018-08-01 00:00:00.000000
                    'new_to': (current['Product Published']
                               + datetime.timedelta(days=self.product_new_threshold_days)).strftime('%Y-%m-%d %H:%M:%S.%f')
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

    def __init__(self):
        Logger.highlight('Processing of concrete products...')
        self.product_concretes_orphaned = {}
        self.product_processed_count = 0
        self.last_concrete = None

    def process(self):
        for product in Processor.products:
            current = Processor.products[product]
            if current['Parent SKU'] != '':  # only true concretes
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
                    self.last_concrete = str(current['Product SKU'])
                    msg = 'paired concrete SKU with parent abstract SKU '+ bold(parent['abstract_sku'])
                    Logger.update(current['Product SKU'], msg)
                except KeyError:
                    self.product_concretes_orphaned[str(current['Product SKU'])] = current
        Logger.highlight('Processing of concrete products completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_concretes[self.last_concrete].keys())
        writer.process()
        writer.write()
        Logger.output('concrete products', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))


class ProductAbstractStore:
    target = 'product_abstract_store.csv'
    stores = ['DE', 'AT', 'US']
    product_abstract_stores = []

    def __init__(self):
        Logger.highlight('Processing of abstract product stores...')
        self.product_processed_count = 0
        self.store_processed_count = 0

    def process(self):
        for product in ProductAbstract.product_abstracts:
            for store in self.stores:
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
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_abstract_stores[0].keys())
        writer.process()
        writer.write()
        Logger.output('abstract product stores', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))


class ProductAttributeKey:
    target = 'product_attribute_key.csv'
    product_attribute_keys = {}

    def __init__(self):
        Logger.highlight('Processing of product attribute keys...')
        self.keys_processed_count = 0

    def process(self):
        self.product_attribute_keys[superatribute] = { 
            'attribute_key': superatribute,
            'is_super': True 
            }
        self.keys_processed_count += 1
        Logger.msg('Processed attribute key '+ bold(superatribute))
        Logger.highlight('Processing of product attribute keys completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_attribute_keys[superatribute].keys())
        writer.process()
        writer.write()
        Logger.output('product attribute keys', self.target)
        Logger.msg('Processed '+bold(str(self.keys_processed_count))+' product attribute keys')


class ProductImage:
    target = 'product_image.csv'
    product_images = []
    pseudo_stores = ['DE', 'US']  # DE: AT + DE, US: EN

    def __init__(self):
        Logger.highlight('Processing of product images...')
        self.products_processed_count = 0
        self.images_processed_count = 0
        self.stores_processed_count = 0
        self.missed_products = []
        self.products = {}
        self.current = {}

    def process_store(self, product):
        for store in self.pseudo_stores:
            image = Processor.products[product]['Featured Image']
            if image != '':
                self.images_processed_count += 1
                if 'abstract_sku' in self.current: 
                    self.product_images.append({
                        'abstract_sku': self.current['abstract_sku'],
                        'concrete_sku': '',
                        'image_set_name': 'default',
                        'external_url_large': Processor.products[product]['Featured Image'],
                        'external_url_small': Processor.products[product]['Featured Image'],
                        'locale': getLocale(store)
                    })
                    msg = 'paired abstract SKU with children image '+bold(Processor.products[product]['Featured Image'])\
                          + ' and store '+bold(store)
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
                    msg = 'paired concrete SKU with children image '+bold(Processor.products[product]['Featured Image'])\
                          + ' and store '+bold(store)
                    Logger.update(product, msg)
            elif image == '':
                self.missed_products.append(product)
                msg = 'Processing of the SKU: '+product+' failed, with message '+italic('no product image defined')
                Logger.warning(msg)
            self.stores_processed_count += 1

    def process(self):
        for product in ProductAbstract.product_abstracts:
            self.current = ProductAbstract.product_abstracts[product]
            self.process_store(product)
        for product in ProductConcrete.product_concretes:
            self.current = ProductConcrete.product_concretes[product]
            self.process_store(product)
        Logger.highlight('Processing of product images keys completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_images[-1])
        writer.process()
        writer.write()
        Logger.output('product images', self.target)
        Logger.msg('Processed '+bold(str(self.products_processed_count))+' SKUs with '
                   + bold(str(self.images_processed_count))+ ' images in '+bold(str(len(self.pseudo_stores)))
                   + ' stores from '+bold(str(len(Processor.products)))+' imported products.')


class ProductImageInternal:
    target = 'product_image_internal.csv'
    product_images = []

    def __init__(self):
        Logger.highlight('Processing of product images internal...')
        self.products_processed_count = 0
        self.images_processed_count = 0

    def process(self):
        self.product_images.append({
            'image_set_name': '',
            'external_url_large': '',
            'external_url_small': '',
            'locale': '',
            'abstract_sku': '',
            'concrete_sku': ''
        })
        Logger.highlight('Processing of product images internal completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_images[-1])
        writer.process()
        writer.write()
        Logger.output('product images internal', self.target)


class ProductLabel:
    target = 'product_label.csv'

    def __init__(self):
        Logger.highlight('Processing of product labels...')

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
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_labels['SALE'])
        writer.process()
        writer.write()
        Logger.output('product labels', self.target)
        Logger.msg('Processed '+bold(str(len(self.product_labels)))+' product labels.')


class ProductManagmentAttribute:
    target = 'product_management_attribute.csv'
    product_management_attributes = []

    def __init__(self):
        self.products_processed_count = 0
        self.attributes_processed_count = 0
        Logger.highlight('Processing of product management attributes...')

    def get_attributes(self):
        attributes = []
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            attribute = current['value_1']
            self.products_processed_count += 1
            if attribute not in attributes:
                attributes.append(attribute)
                self.attributes_processed_count += 1
                Logger.msg('Processing of the SKU: '+bold(product)+' completed with message: '
                           + italic('collected a new attribute: '+bold(attribute)))
        return ','.join(list(OrderedDict.fromkeys(attributes)))

    def process(self):
        attributes = self.get_attributes()
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
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_management_attributes[-1])
        writer.process()
        writer.write()
        Logger.output('product managment attributes', self.target)
        Logger.msg('Processed '+bold(superatribute)+' with values '+bold(attributes)+' from '+bold(self.products_processed_count)+' concrete products.')


class ProductPrice:
    target = 'product_price.csv' 
    product_prices = []

    def __init__(self):
        Logger.highlight('Processing of product prices...')
        self.missed_product_prices = []
        self.products_processed_count = 0
        self.prices_processed_count = 0

    def process(self):
        for product in ProductConcrete.product_concretes:
            current = ProductConcrete.product_concretes[product]
            try: 
                price = float(Processor.products[product]['Price'])
                tax = 0.2 * price
                net = int((price - tax) * 100)
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
                    Logger.warning('Processing of the SKU: '+product+' failed, with message: '
                                   + italic('net price > gross price.'))
                    continue
                self.prices_processed_count += 2
                Logger.msg('Processing of the SKU: '+bold(product)+' completed with the message: '
                           + italic('paired concrete product with net/gross prices: '+bold([net, gross])+' for store: '+bold(store)))
            except KeyError: 
                self.missed_product_prices.append(product)
                Logger.warning('Processing of the SKU: '+product+' failed, with message: '
                               + italic('no product price defined.'))
        Logger.highlight('Processing of product prices completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_prices[-1])
        writer.process()
        writer.write()
        Logger.output('product prices', self.target)
        Logger.msg('Processed ' + bold(self.products_processed_count)+' SKUs with '+bold(self.prices_processed_count)
                   + ' prices (net+gross) for store: '+bold(store))


class ProductStock:
    target = 'product_stock.csv'
    product_stocks = {}

    def __init__(self):
        self.last_product = None
        self.products_processed_count = 0
        Logger.highlight('Processing of product stocks...')

    def process(self):
        for product in ProductConcrete.product_concretes:
            current = Processor.products[product]
            self.product_stocks[product] = {
                'concrete_sku': str(current['Parent SKU']),
                'name': 'Warehouse1',
                'quantity': current['Quantity'],
                'is_never_out_of_stock': False,
                'is_bundle': False
            }
            self.last_product = product
            self.products_processed_count += 1
            Logger.msg('Processing of the SKU: '+bold(product)+' completed with the message: '
                       + italic('paired concrete product with stock value: '+bold(self.product_stocks[product]['quantity'])))
        Logger.highlight('Processing of product stocks completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_stocks[self.last_product])
        writer.process()
        writer.write()
        Logger.output('product stocks', self.target)
        Logger.msg('Processed '+bold(self.products_processed_count)+' SKUs.')


class ProductDiscontinued:
    target = 'product_discontinued.csv'
    products_discontinued = {}

    def __init__(self):
        Logger.highlight('Processing of products discontinued...')
        self.last_product = None
        self.products_processed_count = 0

    def process(self):
        sku = ''
        self.products_discontinued[sku] = {
            'sku_concrete': '',
            'note.en_US': '',
            'note.de_DE': ''
        }
        self.products_processed_count += 1
        Logger.msg('Processing of the SKU: '+bold(sku)+' completed.')
        Logger.highlight('Processing of products discontinued completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.products_discontinued[sku])
        writer.process()
        writer.write()
        Logger.output('products discontinued', self.target)
        Logger.msg('Processed '+bold(self.products_processed_count)+' SKUs.')


class ProductGroup:
    target = 'product_group.csv'
    product_groups = {}

    def __init__(self):
        Logger.highlight('Processing of product groups...')
        self.last_product = None
        self.products_processed_count = 0

    def process(self):
        sku = ''
        self.product_groups[sku] = {
            'group_key': '',
            'abstract_sku': '',
            'position': ''
        }
        self.products_processed_count += 1
        Logger.msg('Processing of the SKU: '+bold(sku)+' completed.')
        Logger.highlight('Processing of product groups completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_groups[sku])
        writer.process()
        writer.write()
        Logger.output('product groups', self.target)
        Logger.msg('Processed '+bold(self.products_processed_count)+' SKUs.')


class ProductSearchAttributeMap:
    target = 'product_search_attribute_map.csv'
    search_attributes = []
    target_fields = ['full-text-boosted', 'suggestion-terms', 'completion-terms']

    def __init__(self):
        Logger.highlight('Processing of product search attribute map...')
        self.attributes_processed_count = 0

    def process(self):
        for target in self.target_fields:
            self.search_attributes.append({
                'attribute_key': superatribute,
                'target_field': target
            })
            self.attributes_processed_count += 1
            Logger.msg('Processing of the attribute: '+bold(superatribute)+' completed.')
        Logger.highlight('Processing of product search attribute map completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.search_attributes[-1])
        writer.process()
        writer.write()
        Logger.output('product search attribute map', self.target)
        Logger.msg('Processed '+bold(self.attributes_processed_count)+' attributes.')


class ProductSearchAttribute:
    target = 'product_search_attribute.csv'
    search_attributes = {}

    def __init__(self):
        Logger.highlight('Processing of product search attributes...')
        self.attributes_processed_count = 0

    def process(self):
        self.search_attributes[superatribute] = {
            'key': superatribute,
            'filter_type': 'multi-select',  # single-select
            'position': 1,
            'key.en_US': superatribute,
            'key.de_DE': superatribute
        }
        Logger.msg('Processing of the attribute: ' + bold(superatribute) + ' completed.')
        Logger.highlight('Processing of product search attributes completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.search_attributes[superatribute])
        writer.process()
        writer.write()
        Logger.output('product search attributes', self.target)
        Logger.msg('Processed '+bold(self.attributes_processed_count)+' attributes.')


class Logger:
    @staticmethod
    def highlight(msg):
        print(highlight_green(msg))

    @staticmethod
    def update(sku, msg=False):
        if msg is False:
            print(green('Processing of the SKU: ' + bold(str(sku)) + ' completed.'))
        else:
            print(green('Processing of the SKU: ' + bold(str(sku)) + ' completed with message: ' + italic(msg)))

    @staticmethod
    def summary(num, total, msg=False):
        if msg is False:
            print(green('Processed ' + bold(str(num)) + ' SKUs of ' + bold(str(total)) + ' imported products.'))
        else:
            print(green('Processed ' + bold(str(num)) + ' SKUs of ' + bold(str(total))
                        + ' imported products with message: ') + italic(msg))

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

    @staticmethod
    def intro():
        # lines = [
        #     u' █████╗ ██╗   ██╗ ██████╗  ██████╗ █████╗ ████████╗ ██████╗ ',
        #     u'██╔══██╗██║   ██║██╔═══██╗██╔════╝██╔══██╗╚══██╔══╝██╔═══██╗',
        #     u'███████║██║   ██║██║   ██║██║     ███████║   ██║   ██║   ██║',
        #     u'██╔══██║╚██╗ ██╔╝██║   ██║██║     ██╔══██║   ██║   ██║   ██║',
        #     u'██║  ██║ ╚████╔╝ ╚██████╔╝╚██████╗██║  ██║   ██║   ╚██████╔╝',
        #     u'╚═╝  ╚═╝  ╚═══╝   ╚═════╝  ╚═════╝╚═╝  ╚═╝   ╚═╝    ╚═════╝ '
        # ]
        print(highlight_green(' [  A  V  O  C  A  T  O  ] '))
        print(highlight_green(' [ Products and meta data processor ] '))
        # for line in lines:
        # current = line.replace('█', '\u2588').replace('╗', '\u2557').replace('╔', '\u2554').\
        #     replace('═', '\u2550').replace('║', '\u2551').replace('╗', '\u2557').replace('╚', '\u255A').\
        #     replace('╝', '\u255D')
        # current = line.decode('utf-8')
        print(green('version: '+__version__))
        print(green('license: '+__license__))
        print(green('author:  '+__author__))


class Writer:
    path = None
    location = None

    def __init__(self, target):
        self.file = target
        self.path = os.path.dirname(os.path.realpath(__file__))
        self.target = target
        self.queue = {}
        self.writer = None
        Logger.highlight('Processing of '+self.target+'...')

    def get_fieldnames(self, headers):
        self.location = self.path+'/'+self.file
        self.file = open(self.location, mode='w')
        self.writer = csv.DictWriter(self.file, fieldnames=headers)
        if len(headers) > 0:
            self.writer.writeheader()
        Logger.msg('Created csv file with headers: '+bold(','.join(headers))+' in location: '+bold(self.location))

    def process(self):
        if self.target == 'category_template.csv':
            self.queue = CategoryTemplate.category_templates
            CategoryTemplate.category_templates = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' category template') + ' entities.')
        if self.target == 'category.csv':
            self.queue = {}
            for level in Category.categories:
                self.queue.update(Category.categories[level])
            Category.categories = {}
            Logger.msg('Collected '+bold(str(len(self.queue))+' category')+' entities.')
        if self.target == 'product_abstract.csv':
            self.queue = ProductAbstract.product_abstracts
            ProductAbstract.product_abstracts = {}
            Logger.msg('Collected '+bold(str(len(self.queue))+' abstract product')+' entities.')
        if self.target == 'product_concrete.csv':
            self.queue = ProductConcrete.product_concretes
            ProductConcrete.product_concretes = {}
            Logger.msg('Collected '+bold(str(len(self.queue))+' concrete product')+' entities.')
        if self.target == 'product_abstract_store.csv':
            self.queue = ProductAbstractStore.product_abstract_stores
            ProductAbstractStore.product_abstract_stores = []
            Logger.msg('Collected '+bold(str(len(self.queue))+' product abstract store')+' entities.')
        if self.target == 'product_attribute_key.csv':
            self.queue = ProductAttributeKey.product_attribute_keys
            ProductAttributeKey.product_attribute_keys = {}
            Logger.msg('Collected '+bold(str(len(self.queue))+' product abstract attribute keys')+' entities.')
        if self.target == 'product_image.csv':
            self.queue = ProductImage.product_images
            ProductImage.product_images = []
            Logger.msg('Collected '+bold(str(len(self.queue))+' product image')+' entities.')
        if self.target == 'product_label.csv':
            self.queue = ProductLabel.product_labels
            ProductLabel.product_labels = {}
            Logger.msg('Collected '+bold(str(len(self.queue))+' product label')+' entities.')
        if self.target == 'product_management_attribute.csv':
            self.queue = ProductManagmentAttribute.product_management_attributes
            ProductManagmentAttribute.product_management_attributes = []
            Logger.msg('Collected '+bold(str(len(self.queue))+' product managment attribute')+' entities.')
        if self.target == 'product_price.csv':
            self.queue = ProductPrice.product_prices
            ProductPrice.product_prices = []
            Logger.msg('Collected '+bold(str(len(self.queue))+' product price')+' entities.')
        if self.target == 'product_stock.csv':
            self.queue = ProductStock.product_stocks
            ProductStock.product_stocks = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' product stock') + ' entities.')
        if self.target == 'product_image_internal.csv':
            self.queue = ProductImageInternal.product_images
            ProductImageInternal.product_images = []
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' product image internal') + ' entities.')
        if self.target == 'product_discontinued.csv':
            self.queue = ProductDiscontinued.products_discontinued
            ProductDiscontinued.products_discontinued = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' product discontinued') + ' entities.')
        if self.target == 'product_group.csv':
            self.queue = ProductGroup.product_groups
            ProductGroup.product_groups = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' product group') + ' entities.')
        if self.target == 'product_search_attribute_map.csv':
            self.queue = ProductSearchAttributeMap.search_attributes
            ProductSearchAttributeMap.search_attributes = []
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' product search attribute map') + ' entities.')
        if self.target == 'product_search_attribute.csv':
            self.queue = ProductSearchAttribute.search_attributes
            ProductSearchAttribute.search_attributes = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' product search attribute') + ' entities.')
        for index, item in enumerate(self.queue):
            queue_type = type(self.queue)
            current_item = None
            if queue_type is dict:
                current_item = self.queue[item]
            if queue_type is list:
                current_item = self.queue[index]
            if current_item is not None:
                for key in current_item:
                    current_value = current_item[key]
                    current_type = type(current_value)
                    if current_type is unicode:
                        new_value = str(current_value.encode('utf-8').replace('_x000D_\n', '').replace('\xc2', '').replace('\xa0', ''))
                        current_item[key] = new_value
                    if current_type is int: continue
                if queue_type is dict:
                    current_item = self.queue[item] = current_item
                    Logger.msg('Processed queue dict item: '+bold(item)+'')
                if queue_type is list:
                    current_item = self.queue[index] = current_item
                    Logger.msg('Processed queue list item: '+bold(index)+'')

    def write(self):
        count = 0
        if self.target == 'category_template.csv':
            for template in self.queue:
                current = self.queue[template]
                self.writer.writerow(current)
                CategoryTemplate.category_templates[template] = current
                self.queue[template] = None
                count += 1
        if self.target == 'category.csv':
            for category in self.queue:
                current = self.queue[category]
                self.writer.writerow(current)
                Category.categories[category] = current
                self.queue[category] = None
                count += 1
        if self.target == 'product_abstract.csv':
            for product in self.queue:
                current = self.queue[product]
                self.writer.writerow(current)
                ProductAbstract.product_abstracts[product] = current
                self.queue[product] = None
                count += 1
        if self.target == 'product_concrete.csv':
            for product in self.queue:
                current = self.queue[product]
                self.writer.writerow(current)
                ProductConcrete.product_concretes[product] = current
                self.queue[product] = None
                count += 1
        if self.target == 'product_abstract_store.csv':
            for index, store in enumerate(self.queue):
                self.writer.writerow(store)
                ProductAbstractStore.product_abstract_stores.append(store)
                self.queue[index] = None
                count += 1
        if self.target == 'product_attribute_key.csv':
            for key in self.queue:
                current = self.queue[key]
                self.writer.writerow(current)
                ProductAttributeKey.product_attribute_keys[key] = current
                self.queue[key] = None
                count += 1
        if self.target == 'product_image.csv':
            for index, image in enumerate(self.queue):
                self.writer.writerow(image)
                ProductImage.product_images.append(image)
                self.queue[index] = None
                count += 1
        if self.target == 'product_label.csv':
            for label in self.queue:
                current = self.queue[label]
                self.writer.writerow(current)
                ProductLabel.product_labels = current
                self.queue[label] = None
        if self.target == 'product_management_attribute.csv':
            for index, attribute in enumerate(self.queue):
                self.writer.writerow(attribute)
                ProductManagmentAttribute.product_management_attributes.append(attribute)
                self.queue[index] = None
                count += 1
        if self.target == 'product_price.csv':
            for index, price in enumerate(self.queue):
                self.writer.writerow(price)
                ProductPrice.product_prices.append(price)
                self.queue[index] = None
                count += 1
        if self.target == 'product_stock.csv':
            for stock in self.queue:
                current = self.queue[stock]
                self.writer.writerow(current)
                ProductStock.product_stocks[stock] = current
                self.queue[stock] = None
                count += 1
        if self.target == 'product_image_internal.csv':
            for index, image in enumerate(self.queue):
                self.writer.writerow(image)
                ProductImageInternal.product_images.append(image)
                self.queue[index] = None
                count += 1
        if self.target == 'product_discontinued.csv':
            for product in self.queue:
                current = self.queue[product]
                self.writer.writerow(current)
                ProductDiscontinued.products_discontinued[product] = current
                self.queue[product] = None
                count += 1
        if self.target == 'product_group.csv':
            for group in self.queue:
                current = self.queue[group]
                self.writer.writerow(current)
                ProductGroup.product_groups[group] = current
                self.queue[group] = None
                count += 1
        if self.target == 'product_search_attribute_map.csv':
            for index, attribute in enumerate(self.queue):
                self.writer.writerow(attribute)
                ProductSearchAttributeMap.search_attributes.append(attribute)
                self.queue[index] = None
                count += 1
        if self.target == 'product_search_attribute.csv':
            for attribute in self.queue:
                current = self.queue[attribute]
                self.writer.writerow(current)
                ProductSearchAttribute.search_attributes[attribute] = current
                self.queue[attribute] = None
        self.file.close()
        self.queue = {}
        Logger.highlight('Saved '+bold(count)+' queue items.')


class Processor:
    products = {}
    categories = []

    def __init__(self, args):
        Logger.highlight('Recived arguments: '+str(args))
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
                if sku != '':
                    self.products[str(product['Product SKU'])] = product
                if sku == '':
                    missed_ids.append(product['Product ID'])
                del current_row
        for product in self.products:
            current = self.products[product]
            split = current['Category'].split('>')
            # if split[0] != 'Uncategorized': category_mappings[product] = { split[0]: split[1] }
            if split[0] != 'Uncategorized':
                self.categories.append({ 'level_1': split[0], 'level_2': split[1] })
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

    @staticmethod
    def run():
        CategoryTemplate().process()
        Category().process()
        ProductAbstract().process()
        ProductConcrete().process()
        ProductAbstractStore().process()
        ProductAttributeKey().process()
        ProductImage().process()
        ProductImageInternal().process()
        ProductLabel().process()
        ProductManagmentAttribute().process()
        ProductPrice().process()
        ProductStock().process()
        ProductDiscontinued().process()
        ProductGroup().process()
        ProductSearchAttributeMap().process()
        ProductSearchAttribute().process()
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
        return {'headers': headers, 'rows': values}


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
    Logger().intro()
    processor = Processor(args)
    processor.hydrate()
    processor.run()

