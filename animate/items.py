# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://docs.scrapy.org/en/latest/topics/items.html

import scrapy


class AnimateItem(scrapy.Item):
    datetime = scrapy.Field()
    name = scrapy.Field()
    sku = scrapy.Field()
    point = scrapy.Field()
    price = scrapy.Field()
    url = scrapy.Field()
    description = scrapy.Field()
    sale_start = scrapy.Field()
    sale_status = scrapy.Field()
    bonus = scrapy.Field()
    site = scrapy.Field()
    purchase_limit = scrapy.Field()
    genre = scrapy.Field()
    image_urls = scrapy.Field()
    thumbnail_path = scrapy.Field()
    
    
