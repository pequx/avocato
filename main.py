# -*- coding: utf-8 -*-
"""
Module Docstring
"""

__author__ = "https://github.com/pequx"
__version__ = "0.6.0"
__license__ = "Proprietary"

import re, datetime
from os import path as os_path
from shutil import move as os_move
from csv import DictWriter
from argparse import ArgumentParser
from openpyxl import load_workbook
from collections import OrderedDict
from fabulous.color import bold, highlight_red, highlight_green, green, italic, highlight_yellow
from inflection import titleize, pluralize, parameterize, dasherize


superattribute = 'superattribute'
spryker_path = None


class CategoryTemplate:
    target = 'category_template.csv'
    category_templates = {}
    templates_processed_count = 0

    def __init__(self):
        Logger.highlight('Processing of category templates...')

    def process(self):
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
        Logger.highlight('Processing of category templates completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.category_templates['Catalog (default)'].keys())
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output('category templates', self.target)


class Category:
    target = 'category.csv'
    categories = {
        'level_1': {},
        'level_2': {}
    }

    def __init__(self):
        Logger.highlight('Processing of categories...')
        self.mappings_processed_count = 0
        self.categories_processed_count = {
            'level_1': 0,
            'level_2': 0
        }
        self.node_order = {
            'level_1': 10,
            'level_2': 10
        }
        self.last_category = None
        self.categories['level_1']['root'] = {
            'category_key': 'demoshop',
            'parent_category_key': '',
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
            'node_order': '',
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
            'node_order': 30,
            'template_name': 'Catalog (default)'
        }
        Logger.msg('Processed two system categories: '+bold('root,product-bundle'))

    def process(self):
        for mapping in Processor.categories:
            level_1 = str(mapping['level_1'].lower().replace(' ', '-'))
            if level_1 not in self.categories['level_1']:
                name = {
                    'de_DE': titleize(mapping['level_1']),
                    'en_US': titleize(mapping['level_1'])
                }
                keywords = {
                    'de_DE': ', '.join(name['de_DE'].split(' ')),
                    'en_US': ', '.join(name['en_US'].split(' '))
                }
                self.categories['level_1'][level_1] = {
                    'category_key': level_1,
                    'parent_category_key': 'demoshop',
                    'name.de_DE': name['de_DE'],
                    'name.en_US': name['en_US'],
                    'meta_title.de_DE': name['de_DE'],
                    'meta_title.en_US': name['de_DE'],
                    'meta_description.de_DE': name['de_DE'],
                    'meta_description.en_US': name['de_DE'],
                    'meta_keywords.de_DE': keywords['de_DE'],
                    'meta_keywords.en_US': keywords['en_US'],
                    'is_active': 1,
                    'is_in_menu': 1,
                    'is_clickable': 1,
                    'is_searchable': 1,
                    'is_root': 0,
                    'is_main': 1,
                    'node_order': self.node_order['level_1'],
                    'template_name': 'Catalog + CMS Block'
                }
                self.categories_processed_count['level_1'] += 1
                self.node_order['level_1'] += 10
                self.last_category = level_1
                Logger.msg('Processing of the category level 1: '+bold(level_1)+' completed with parent: '+bold('demoshop'))
            level_2 = str(mapping['level_2'].lower().replace(' ', '-'))
            if level_2 not in self.categories['level_2']:
                name = {
                    'de_DE': titleize(mapping['level_2']),
                    'en_US': titleize(mapping['level_2'])
                }
                keywords = {
                    'de_DE': ', '.join(name['de_DE'].split(' ')),
                    'en_US': ', '.join(name['en_US'].split(' '))
                }
                # level_2 = level_2+'-'+str(uuid.uuid4())
                self.categories['level_2'][level_2] = {
                    'category_key': lowerCase(mapping['level_2']),
                    'parent_category_key': level_1,
                    'name.de_DE': name['de_DE'],
                    'name.en_US': name['en_US'],
                    'meta_title.de_DE': name['de_DE'],
                    'meta_title.en_US': name['de_DE'],
                    'meta_description.de_DE': name['de_DE'],
                    'meta_description.en_US': name['de_DE'],
                    'meta_keywords.de_DE': keywords['de_DE'],
                    'meta_keywords.en_US': keywords['en_US'],
                    'is_active': 1,
                    'is_in_menu': 1,
                    'is_clickable': 1,
                    'is_searchable': 1,
                    'is_root': 0,
                    'is_main': 0,
                    'node_order': 10,
                    'template_name': 'Catalog + CMS Block'
                }
                self.categories_processed_count['level_2'] += 1
                self.node_order['level_2'] += 10
                # if mapping['Category'] != 'Private':
                Logger.msg('Processing of the category level 2: '+bold(level_2)+' completed with parent: '+bold(level_1))
                self.mappings_processed_count += 1
        Logger.highlight('Processing of categories completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.categories['level_1'][self.last_category].keys())
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output('categories', self.target)
        Logger.msg('Processed '+bold(str(self.mappings_processed_count))
                   + ' category mappings providing '+bold(str(self.categories_processed_count['level_1']))
                   + ' level 1 categories, and '+bold(str(self.categories_processed_count['level_2']))
                   + ' level 2 categories.')


class CmsBlockCategoryPosition:
    name = 'CMS Block Category Position'
    target = 'cms_block_category_position.csv'
    cms_block_category_positions = {}

    def __init__(self):
        Logger.highlight('Processing of the '+self.name+'s is starting...')
        self.positions = ['top', 'middle', 'bottom']
        self.processed_positions_count = 0

    def process(self):
        for position in self.positions:
            self.cms_block_category_positions[position] = {
                'cms_block_category_position_name': position
            }
            self.processed_positions_count += 1
            Logger.msg('Processing of the position ' + bold(position) + ' completed.')
        if self.processed_positions_count == 3:
            Logger.highlight('Processing of the '+self.name+'s is completed.')
            writer = Writer(self.target)
            writer.get_fieldnames(self.cms_block_category_positions['bottom'].keys())
            writer.process()
            writer.write()
            connector = Connector(self.target)
            connector.connect()
            Logger.output(self.name, self.target)
            Logger.msg('Processed '+bold(str(self.processed_positions_count))+' of '+self.name+'s.')


class CmsBlockStore:
    name = 'CMS Block Store'
    target = 'cms_block_store.csv'
    cms_block_stores = []
    processed_count = {
        'store': 0,
        'block': 0
    }

    def __init__(self, target, stores):
        Logger.highlight('Processing of the '+self.name+' is starting...')
        self.target = target  # cms_block_store.csv
        self.stores = stores  # ['DE', 'AT', 'US']
        self.blocks = [
            'Home Page',
            'Teaser for home page',
            'Product SEO content',
            'Category CMS page showcase for Top position',
            'Category CMS page showcase for Middle position',
            'Category CMS page showcase for Bottom position',
            'CMS block for category Computers',
            'Main slide-1',
            'Main slide-2',
            'Main slide-3',
            'Featured Products',
            'Top Sellers',
            'Featured Categories',
            'Category Banner-1',
            'Category Banner-2',
            'Category Banner-3',
            'Category Banner-4',
            'Product CMS Block',
            'Category Block Bottom',
            'Category Block Middle'
        ]

    def process(self):
        for store in self.stores:
            for block in self.blocks:
                self.cms_block_stores.append({
                    'block_name': block,
                    'store_name': store
                })
                self.processed_count['block'] += 1
                Logger.msg('Processing of the block ' + bold(block) + ' completed.')
            self.processed_count['store'] += 1
            Logger.msg('Processing of the store ' + bold(store) + ' completed.')
        Logger.highlight('Processing of the' + self.name+' is completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.cms_block_stores[-1])
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output(self.name, self.target)


class Navigation:
    name = 'Navigation'
    target = 'navigation.csv'
    navigation_items = {}

    def __init__(self, keys):
        Logger.highlight('Processing of the ' + self.name + ' is starting...')
        self.keys = keys
        self.processed_count = {
            'key': 0
        }
        self.last_key = None

    def process(self):
        for key in self.keys:
            self.navigation_items[key] = {
                'key': key,
                'name': titleize('key')
            }
            self.processed_count['key'] += 1
            self.last_key = key
            Logger.msg('Processing of the navigation key ' + bold(key) + ' is completed.')
        Logger.highlight('Processing of the ' + self.name + ' is completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.navigation_items[self.last_key].keys())
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output(self.name, self.target)


class NavigationNode:
    name = 'Navigation Node'
    target = 'navigation_node.csv'
    navigation_nodes = {}

    def __init__(self, nodes):
        Logger.highlight('Processing of the '+pluralize(self.name)+' is starting...')
        self.processed_count = {
            'node': 0,
            'category': 0
        }
        self.missed = {
            'category': []
        }
        self.prefix = {
            'category': '/product/'
        }
        self.nodes = nodes
        self.last_node = None

    def get_node_key(self):
        return 'node_key_'+str(len(self.navigation_nodes) + 1)

    def add_nodes(self):
        """Adds nodes from the init"""
        for index, node in enumerate(self.nodes):
            key = self.get_node_key()
            node['node_key'] = key
            self.navigation_nodes[key] = node
            self.processed_count['node'] += 1
            Logger.msg('Added extra node '+bold(node['title.en_US']))
        del self.nodes

    def get_category_nodes(self):
        """Provides category nodes mapping"""
        for category in Category.categories:
            try:
                current = Category.categories[category]
                is_valid = current['is_active'] == 1 and current['is_in_menu'] == 1 and current['is_root'] == 0
                # is_level_1 = current['is_main'] == 1 and \
                #              current['parent_category_key'] == root_node['category_key']
                # is_level_2 = current['is_main'] == 0 and current['parent_category_key'] != ''
                if is_valid:
                    # url_param = current['meta_title.en_US'].lower().replace(' ', '-')
                    node_key = self.get_node_key()
                    self.navigation_nodes[node_key] = {
                        'navigation_key': 'MAIN_NAVIGATION',
                        'node_key': node_key,
                        'parent_node_key': None,
                        'node_type': 'category',
                        'title.en_US': current['meta_title.en_US'],
                        'url.en_US': '/en/'+category+'/',
                        'css_class.en_US': 'new-color',
                        'title.de_DE': current['meta_title.de_DE'],
                        'url.de_DE': '/de/'+category+'/',
                        'css_class.de_DE': 'new-color',
                        'valid_from': '',
                        'valid_to': ''
                    }
                    # self.navigation_nodes[node_key]['parent_node_key'] = current['parent_category_key']
                    self.processed_count['node'] += 1
                    self.processed_count['category'] += 1
                    self.last_node = node_key
                    Logger.msg('Processing of the '+bold(current['category_key'])
                               + ' category node is completed, message:'
                               + italic(' acceptance criteria: '+str(is_valid)+','))
            except KeyError:
                # self.missed_count['category'] += 1
                # self.missed_count['node'] += 1
                self.missed['category'].append(category)
                Logger.warning('Missed category '+bold(category)+'.')
        Logger.msg('Processing of the '+bold(str(self.processed_count['category']))+' categories is completed.')
        Logger.msg('Missed categories count: '+bold(str(len(self.missed['category']))))

    def get_node_parents(self):
        parents = {}
        for node in self.navigation_nodes:
            try:
                current_navigation_node = self.navigation_nodes[node]
                current_node_key = current_navigation_node['title.de_DE'].lower().replace(' ', '-')
                current_category = Category.categories[current_node_key]
                current_parent = Category.categories[current_category['parent_category_key']]
                parents[node] = current_parent['meta_description.en_US']
                print 'test'
            except KeyError:
                continue
        for parent_node in parents:
            for node in self.navigation_nodes:
                current_node = self.navigation_nodes[node]
                if current_node['title.en_US'] == parents[parent_node]:
                    node_match = self.navigation_nodes[node]
                    current_node['parent_node_key'] = parent_node
                    parents[parent_node] = node_match['node_key']
                    print 'test'
                print 'ja'
                self.navigation_nodes[parent_node]['parent_node_key'] = parents[parent_node]
            print 'test'
        print 'ja'

    def add_input_nodes(self):
        for node in Processor.inputs['navigation_node']:
            node_key = self.get_node_key()
            current = Processor.inputs['navigation_node'][node]
            current['node_key'] = node_key
            self.navigation_nodes[node_key] = current
            self.processed_count['node'] += 1
            Logger.msg('Processing of the '+bold(current['title.en_US'])+' input node is completed.')

    def process(self):
        self.get_category_nodes()
        self.add_input_nodes()
        self.add_nodes()  # extra nodes from input
        self.get_node_parents()
        Logger.highlight('Processing of the '+pluralize(self.name)+' is completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.navigation_nodes[self.last_node].keys())
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output(self.name, self.target)
        Logger.msg('Processed '+bold(str(self.processed_count['node']))+' nodes.')


class CmsBlockCategory:
    name = 'CMS Block Categories'
    target = 'cms_block_category.csv'
    cms_block_categories = []

    def __init__(self):
        Logger.highlight('Processing of the '+self.name+' is starting...')
        self.processed_categories_count = 0
        self.missed_categories = {}
        self.positions = {
            'bottom': {'block_name': 'Category Block Bottom'},
            'middle': {'block_name': 'Category Block Middle'}
        }

    def process(self):
        for position in self.positions:
            current_position = self.positions[position]
            Logger.msg('Processing of the '+bold(position)+' position...')
            for category in Category.categories:
                try:
                    current_category = Category.categories[category]
                    # Logger.msg('Processing of the '+bold(category)+' category...')
                    self.cms_block_categories.append({
                        'block_name': current_position['block_name'],
                        'category_key': current_category['category_key'],
                        'template_name': current_category['template_name'],
                        'cms_block_category_position_name': position
                    })
                    self.processed_categories_count += 1
                    Logger.msg('Processing of the cms block category position '+bold(str(position))
                               + ' with the category '+bold(str(category))+' is completed.')
                except KeyError:
                    self.missed_categories[category] = Category.categories[category]
                    Logger.warning('Processing of the category key '
                                   + bold(str(category)+' has failed.'))
            Logger.msg('Processing of the ' + bold(position) + ' is completed.')
        Logger.highlight('Processing of the '+self.name+' is completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.cms_block_categories[-1])
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output(self.name, self.target)
        Logger.msg('Processed '+bold(str(self.processed_categories_count))+' of '+self.name+'s.')


class ProductAbstract:
    target = 'product_abstract.csv'
    product_new_threshold_days = 31  # number of days after which product will not be marked as `new`
    product_abstracts = {}

    def __init__(self):
        Logger.highlight('Processing of abstract products...')
        self.product_processed_count = 0
        self.product_missed = {}
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
            return '/de' + match[0]
        if iso == 'en_US':
            return '/en' + match[0]

    def process(self):
        for product in Processor.products:
            current = Processor.products[product]
            # if current['Parent ID'] == '':  # only true abstracts
            category = self.category_key(current['Category'].split('>'))
            sku = str(current['Product SKU'])
            try:
                categories = Category.categories
                category_match = Category.categories[category]
                if current['Parent SKU'] == '':
                    self.product_abstracts[sku] = {
                        'category_key': category_match['category_key'],
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
                        'meta_keywords.de_DE': current['Slug'].replace('-', ' '),
                        'meta_keywords.en_US': current['Slug'].replace('-', ' '),
                        'meta_description.de_DE': current['Description'],
                        'meta_description.en_US': current['Description'],
                        'is_featured': self.is_featured(current['Featured']),
                        'color_code': '#FFFFFF',
                        'new_from': current['Product Published'].strftime('%Y-%m-%d %H:%M:%S.%f'),
                        'new_to': (current['Product Published']
                                   + datetime.timedelta(days=self.product_new_threshold_days)).strftime(
                            '%Y-%m-%d %H:%M:%S.%f')
                    }
                    self.product_processed_count += 1
                    self.last_abstract = str(current['Product SKU'])
                    Logger.update(current['Product SKU'])
            except KeyError:
                self.product_missed[sku] = current
                Logger.warning('Missed SKU: '+bold(sku))
        Logger.highlight('Processing of abstract products completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_abstracts[self.last_abstract].keys())
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
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
            sku = str(current['Product SKU'])
            if current['Parent SKU'] != '':  # only true concretes
                try:
                    parent = ProductAbstract.product_abstracts[str(current['Parent SKU'])]
                    self.product_concretes[sku] = {
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
                        'attribute_key_1': superattribute,
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
        connector = Connector(self.target)
        connector.connect()
        Logger.output('concrete products', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))


class ProductAbstractStore:
    target = 'product_abstract_store.csv'
    product_abstract_stores = []

    def __init__(self):
        Logger.highlight('Processing of abstract product stores...')
        self.product_processed_count = 0
        self.store_processed_count = 0
        self.stores = ['DE', 'AT', 'US']

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
        connector = Connector(self.target)
        connector.connect()
        Logger.output('abstract product stores', self.target)
        Logger.summary(self.product_processed_count, len(Processor.products))


class ProductAttributeKey:
    target = 'product_attribute_key.csv'
    product_attribute_keys = {}

    def __init__(self):
        Logger.highlight('Processing of product attribute keys...')
        self.keys_processed_count = 0

    def process(self):
        self.product_attribute_keys[superattribute] = { 
            'attribute_key': superattribute,
            'is_super': True 
            }
        self.keys_processed_count += 1
        Logger.msg('Processed attribute key '+ bold(superattribute))
        Logger.highlight('Processing of product attribute keys completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.product_attribute_keys[superattribute].keys())
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
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
                    image = Processor.products[product]['Featured Image']
                    self.product_images.append({
                        'abstract_sku': '',
                        'concrete_sku': self.current['concrete_sku'],
                        'image_set_name': 'default',
                        'external_url_large': image,
                        'external_url_small': image,
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
                try:
                    parent_sku = Processor.products[product]['Parent SKU']
                    parent = Processor.products[str(parent_sku)]
                    self.product_images.append({
                        'abstract_sku': '',
                        'concrete_sku': self.current['concrete_sku'],
                        'image_set_name': 'default',
                        'external_url_large': parent['Featured Image'],
                        'external_url_small': parent['Featured Image'],
                        'locale': getLocale(store)
                    })
                    self.products_processed_count += 1
                    msg = 'automatically paired concrete SKU with children image '+bold(Processor.products[product]['Featured Image'])\
                          + ' and store '+bold(store)
                    Logger.update(product, msg)
                except KeyError:
                    print 'ja'
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
        connector = Connector(self.target)
        connector.connect()
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
        connector = Connector(self.target)
        connector.connect()
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
        connector = Connector(self.target)
        connector.connect()
        Logger.output('product labels', self.target)
        Logger.msg('Processed '+bold(str(len(self.product_labels)))+' product labels.') 


class ProductManagementAttribute:
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
            'key': superattribute,
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
        connector = Connector(self.target)
        connector.connect()
        Logger.output('product managment attributes', self.target)
        Logger.msg('Processed '+bold(superattribute)+' with values '+bold(attributes)+' from '+bold(self.products_processed_count)+' concrete products.')


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
        connector = Connector(self.target)
        connector.connect()
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
        for product in Processor.products:
            current = Processor.products[product]
            if current['Quantity'] != '':
                self.product_stocks[product] = {
                    'concrete_sku': current['Product SKU'],
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
        connector = Connector(self.target)
        connector.connect()
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
        connector = Connector(self.target)
        connector.connect()
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
        connector = Connector(self.target)
        connector.connect()
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
                'attribute_key': superattribute,
                'target_field': target
            })
            self.attributes_processed_count += 1
            Logger.msg('Processing of the attribute: '+bold(superattribute)+' completed.')
        Logger.highlight('Processing of product search attribute map completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.search_attributes[-1])
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
        Logger.output('product search attribute map', self.target)
        Logger.msg('Processed '+bold(self.attributes_processed_count)+' attributes.')


class ProductSearchAttribute:
    target = 'product_search_attribute.csv'
    search_attributes = {}

    def __init__(self):
        Logger.highlight('Processing of product search attributes...')
        self.attributes_processed_count = 0

    def process(self):
        self.search_attributes[superattribute] = {
            'key': superattribute,
            'filter_type': 'multi-select',  # single-select
            'position': 1,
            'key.en_US': superattribute,
            'key.de_DE': superattribute
        }
        Logger.msg('Processing of the attribute: ' + bold(superattribute) + ' completed.')
        Logger.highlight('Processing of product search attributes completed.')
        writer = Writer(self.target)
        writer.get_fieldnames(self.search_attributes[superattribute])
        writer.process()
        writer.write()
        connector = Connector(self.target)
        connector.connect()
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
        path = os_path.dirname(os_path.realpath(__file__))
        print(green('Finished processing of '+bold(type)))

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
        print(highlight_green(' [ Products and meta data processor for Spryker OS ]'))
        # for line in lines:
        # current = line.replace('█', '\u2588').replace('╗', '\u2557').replace('╔', '\u2554').\
        #     replace('═', '\u2550').replace('║', '\u2551').replace('╗', '\u2557').replace('╚', '\u255A').\
        #     replace('╝', '\u255D')
        # current = line.decode('utf-8')
        print(green('\tversion: ')+__version__)
        print(green('\tlicense: ')+__license__)
        print(green('\tauthor:  ')+__author__)


class Writer:
    path = None
    location = None

    def __init__(self, target):
        self.file = target
        self.path = os_path.dirname(os_path.realpath(__file__))
        self.target = target
        self.queue = {}
        self.writer = None
        self.process_failures = {}
        Logger.highlight('Processing of '+self.target+'...')

    def get_fieldnames(self, headers):
        self.location = self.path+'/'+self.file
        self.file = open(self.location, mode='w')
        self.writer = DictWriter(self.file, fieldnames=headers)
        if len(headers) > 0:
            self.writer.writeheader()
        Logger.msg('Created csv object with headers: '+bold(','.join(headers))+' which will be saved in: '+bold(self.location))

    def process(self):
        if self.target == 'category_template.csv':
            self.queue = CategoryTemplate.category_templates
            CategoryTemplate.category_templates = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' category_template') + ' entities.')
        if self.target == 'category.csv':
            self.queue = {}
            for level in Category.categories:
                for category in Category.categories[level]:
                    current_category = Category.categories[level][category]
                    self.queue[category] = current_category
                # self.queue.update(Category.categories[level])
            Category.categories = {}
            Logger.msg('Collected '+bold(str(len(self.queue))+' category')+' entities.')
        if self.target == 'cms_block_category_position.csv':
            load = CmsBlockCategoryPosition.cms_block_category_positions
            if len(load) < 1:
                self.process_failures[self.target] = {'load': load}
                Logger.warning('File '+self.target+' may be corrupted.')
            self.queue = CmsBlockCategoryPosition.cms_block_category_positions
            CmsBlockCategoryPosition.cms_block_category_positions = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' cms_block_category_position') + ' entities.')
        if self.target == 'cms_block_store.csv':
            self.queue = CmsBlockStore.cms_block_stores
            CmsBlockStore.cms_block_stores = []
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' cms_block_store') + ' entities.')
        if self.target == 'navigation.csv':
            self.queue = Navigation.navigation_items
            Navigation.navigation_items = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' navigation') + ' entities.')
        if self.target == 'navigation_node.csv':
            self.queue = NavigationNode.navigation_nodes
            NavigationNode.navigation_nodes = {}
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' navigation_node') + ' entities.')
        if self.target == 'cms_block_category.csv':
            self.queue = CmsBlockCategory.cms_block_categories
            CmsBlockCategory.cms_block_categories = []
            Logger.msg('Collected ' + bold(str(len(self.queue)) + ' cms_block_category') + ' entities.')
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
            self.queue = ProductManagementAttribute.product_management_attributes
            ProductManagementAttribute.product_management_attributes = []
            Logger.msg('Collected '+bold(str(len(self.queue))+' product management attribute')+' entities.')
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
        if self.target == 'cms_block_category_position.csv':
            for position in self.queue:
                current = self.queue[position]
                self.writer.writerow(current)
                CmsBlockCategoryPosition.cms_block_category_positions[position] = current
                self.queue[position] = None
                count += 1
        if self.target == 'cms_block_store.csv':
            for index, store in enumerate(self.queue):
                self.writer.writerow(store)
                CmsBlockStore.cms_block_stores.append(store)
                self.queue[index] = None
                count += 1
        if self.target == 'navigation.csv':
            for nav_item in self.queue:
                current = self.queue[nav_item]
                self.writer.writerow(current)
                Navigation.navigation_items[nav_item] = current
                self.queue[nav_item] = None
                count += 1
        if self.target == 'navigation_node.csv':
            for node in self.queue:
                current = self.queue[node]
                self.writer.writerow(current)
                NavigationNode.navigation_nodes[node] = current
                self.queue[node] = None
        if self.target == 'cms_block_category.csv':
            for index, category in enumerate(self.queue):
                self.writer.writerow(category)
                CmsBlockCategory.cms_block_categories.append(category)
                self.queue[index] = None
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
                ProductManagementAttribute.product_management_attributes.append(attribute)
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


class Connector:
    def __init__(self, target):
        self.export_path = spryker_path+'/data/import/'+target
        self.import_path = os_path.dirname(__file__)+'/'+target

    def connect(self):
        os_move(self.import_path, self.export_path)
        Logger.msg('Uploaded saved file to '+bold(self.export_path))


class Processor:
    products = {}
    categories = []
    inputs = {}

    def __init__(self, args):
        Logger.msg('Received arguments: '+str(args))
        self.args = args
        workbook = load_workbook(filename=self.args.filename, read_only=self.args.read_only)
        # spryker_path = self.args.spryker_path
        self.data_product_export = self.process_workbook(workbook['Product Export'])
        self.data_product_meta = self.process_workbook(workbook['Product Meta Data'])
        self.inputs['navigation_node'] = self.process_workbook(workbook['navigation_node'])
        del workbook

    def hydrate_input(self):
        Logger.highlight('Hydration of input data sheets...')
        for data in self.inputs:
            current_input = self.inputs[data]
            self.inputs[data] = {}
            for section in current_input:
                current_section = current_input[section]
                if section == 'rows':
                    for row in current_section:
                        result = {}
                        current_row = current_section[row]
                        for index, column in enumerate(current_row):
                            header = str(current_input['headers'][index])
                            # self.inputs[data][index] = {}
                            result[header] = column
                        self.inputs[data][row-2] = result  # Excel gives rows starting form 1=header
                        Logger().msg('Processed row: ' + bold(row))
            Logger.highlight('Hydration of provided input data sheets completed.')

    def hydrate(self):
        missed_ids = []  # products without id
        for row in self.data_product_export['rows']:
            if row > 1:
                current_row = self.data_product_export['rows'][row]
                product = {}
                for index, value in enumerate(current_row):
                    header = self.data_product_export['headers'][index]
                    product[header] = value
                sku = product['Product SKU']
                if sku != '':
                    self.products[str(product['Product SKU'])] = product
                if sku == '':
                    missed_ids.append(product['Product ID'])
                del current_row
        for product in self.products:
            current = self.products[product]
            split = current['Category'].split('>')
            if split[0] != 'Uncategorized':
                self.categories.append({ 'level_1': split[0], 'level_2': split[1] })
            del current
        category_count = {}
        category_index = 0
        while category_index < len(self.categories):
            if self.categories[category_index]['level_2'] in category_count:
                del self.categories[category_index]
            else:
                category_count[self.categories[category_index]['level_2']] = 1
                category_index += 1
        del category_count, category_index
        missed_rows = {}  # rows without data
        missed_matches = []  # meta prodocuts not mached to export products
        for row in self.data_product_meta['rows']:
            if row > 1:
                current_row = self.data_product_meta['rows'][row]
                meta_product = {}
                for index, value in enumerate(current_row):
                    meta_product[self.data_product_meta['headers'][index]] = value
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

    @staticmethod
    def run():
        CategoryTemplate().process()
        Category().process()
        CmsBlockCategoryPosition().process()
        CmsBlockStore(target='cms_block_store.csv', stores=['DE', 'AT', 'US']).process()
        Navigation(keys=[
            'MAIN_NAVIGATION',
            'FOOTER_NAVIGATION',
            'PAYMENT_PROVIDERS',
            'SHIPMENT_PROVIDERS',
            'SOCIAL_LINKS',
            'FOOTER_NAVIGATION_TOP_CATEGORIES',
            'FOOTER_NAVIGATION_POPULAR_BRANDS'
        ]).process()
        NavigationNode(nodes=[
            {
                'navigation_key': None,
                'node_key': None,
                'parent_node_key': None,
                'node_type': 'link',
                'title.en_US': 'New',
                'url.en_US': '/en/new',
                'css_class.en_US': 'new-color',
                'title.de_DE': 'Neu',
                'url.de_DE': '/de/new',
                'css_class.de_DE': 'new-color',
                'valid_from': '',
                'valid_to': ''
            },
            {
                'navigation_key': None,
                'node_key': None,
                'parent_node_key': None,
                'node_type': 'link',
                'title.en_US': 'Sale %',
                'url.en_US': '/en/outlet',
                'css_class.en_US': 'sale-color',
                'title.de_DE': 'Sale %',
                'url.de_DE': '/de/outlet',
                'css_class.de_DE': 'sale-color',
                'valid_from': '',
                'valid_to': ''
            },
        ]).process()
        CmsBlockCategory().process()
        ProductAbstract().process()
        ProductConcrete().process()
        ProductAbstractStore().process()
        ProductAttributeKey().process()
        ProductImage().process()
        ProductImageInternal().process()
        ProductLabel().process()
        ProductManagementAttribute().process()
        ProductPrice().process()
        ProductStock().process()
        ProductDiscontinued().process()
        ProductGroup().process()
        ProductSearchAttributeMap().process()
        ProductSearchAttribute().process()
        Logger.highlight('Processing complete.')

    @staticmethod
    def process_workbook(sheet):
        row_count = 0
        cell_count = 0
        values = {}
        for row in sheet.rows:
            cols = []
            for cell in row:
                if cell.value is not None:
                    cols.append(cell.value)
                elif cell.value is None:
                    cols.append('')
                cell_count += 1
                del cell
            row_count += 1
            values[row_count] = cols
            del cols, row
        headers = values[1]
        del values[1]  # removes first row as it contains headers
        return {'headers': headers, 'rows': values}


def upperCase(string):
    output = string.replace('-', ' ').upper()
    return output


def camelCase(string, space=False):
    output = ''.join(x for x in string.title() if x.isalnum())
    if space: return output[0].lower() + ' ' + output[1:]
    return output[0].lower() + output[1:]


def lowerCase(string):
    output = string.replace(' ', '-').lower()
    return output


def getLocale(store):
    if store == 'DE':
        return 'de_DE'
    if store == 'US':
        return 'en_US'
    if store == 'AT':
        return 'de_DE'


if __name__ == "__main__":
    """ This is executed when run from the command line """
    parser = ArgumentParser()
    # Required positional argument
    parser.add_argument("filename", help="Required positional argument")
    parser.add_argument('-s', '--spryker-path', action='store', dest='spryker_path', help='Path to spryker mount point.')
    parser.add_argument("-r", "--read-only", action="store", default=True, help='Opens the Excel workbook in read-only mode.')
    parser.add_argument('-m', '--memory', action='store', dest='opt_memory', help='Simple garbage collector, removes not neede stuff.')
    parser.add_argument(
        "-v",
        "--verbose",
        action="count",
        default=0,
        help="Verbosity (-v, -vv, etc)")
    parser.add_argument(
        "--version",
        action="version",
        version="%(prog)s (version {version})".format(version=__version__))
    args = parser.parse_args()
    spryker_path = args.spryker_path
    Logger().intro()
    processor = Processor(args)
    processor.hydrate()
    processor.hydrate_input()
    processor.run()

