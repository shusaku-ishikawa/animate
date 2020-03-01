# -*- coding: utf-8 -*-
import scrapy
import re
from ..items import AnimateItem
from datetime import datetime

class ItemSpider(scrapy.Spider):
    name = 'item'
    allowed_domains = ['animate-onlineshop.jp']
    base_url = 'https://www.animate-onlineshop.jp'
    search_page_base_url = f'{base_url}/products/list.php?'

    def extract_sku(self, url):
        m = re.search(r'(?<=\/pd\/)\d+', url)
        return m.group() if m else None

    def extract_price(self, price_str):
        m = re.search(r'(\d|,)+(?=円)', price_str)
        return int(m.group().replace(',', '')) if m else None

    def extract_bonus(self, bonus):
        m = re.search(f'(?<=特典：).+', bonus)
        return m.group().replace('\xa0', '') if m else None

    def start_requests(self):
        for query in self.query_manager.queries:
            url = f'{self.search_page_base_url}{query.querystring}'
            yield scrapy.Request(url, callback=self.parse_result_page, dont_filter = True, meta={'query': query, 'page': 1})

    def parse_result_page(self, response):
        query = response.meta.get('query')
        page = response.meta.get('page')

        item_list = response.xpath('//div[@class="item_list"]/ul/li')
        for item in item_list:
            link_to_item_page = item.xpath('.//div[@class="item_list_thumb"]/a/@href').get()
            link_to_item_page_full = f'{self.base_url}{link_to_item_page}'

            saleprice = response.xpath('//p[@class="saleprice"]/font/text()').get()
            regular_price = item.xpath('.//p[@class="price"]/font/text()').get()
            if saleprice:
                price = int(saleprice.replace(',', ''))
            elif regular_price:
                price = int(regular_price.replace(',', ''))
            sale_status = item.xpath('.//p[@class="stock"]/span/text()').get()
            category_str = item.xpath('.//p[@class="media"]/a/text()').get()
            sale_start = item.xpath('.//p[@class="release"]/text()').get()
            bonus = item.xpath('.//div[@class="item_list_class"]/p/span/text()').get()
            
            if not query.check(price, sale_status):
                #self.logger.info(f'Skipped:{link_to_item_page_full}')
                continue
            yield scrapy.Request(link_to_item_page_full, callback=self.parse_item_page,dont_filter=True, meta={'query': query})
        next_link = [l.xpath('./@href').get() for l in response.xpath('//*[@id="wrapper"]/div[2]/div[2]/section[1]/div[3]/div[1]/div/a') if l.xpath('./text()').get() == '次へ>>']
        if next_link:
            yield scrapy.Request(f'{self.base_url}{next_link[0]}', callback=self.parse_result_page, dont_filter=True, meta={'query': query, 'page': page + 1})

    def parse_item_page(self, response):
        genre = '\n'.join([g for g in response.xpath('//div[@id="breadcrumb"]/ul/li/a/text()').getall() if g != 'ホーム'])
        sku = self.extract_sku(response.url)
        name = response.xpath('//div[@class="item_overview_detail"]/h1/text()').get()
        
        price = self.extract_price(response.xpath('//div[@class="item_price"]/div/p/text()').extract_first())
        point = response.xpath('//*[@id="container"]/div/div[1]/div/section/div[2]/div[1]/div/p[last()]/span[1]/text()').get()
    
        release_date = response.xpath('//*[@id="container"]/div/div[1]/div/section/div[2]/div[2]/p/span/text()').get()
        stock = response.xpath('//*[@id="container"]/div/div[1]/div/section/div[2]/div[2]/div/p[1]/span/text()').get()

        bonus = response.xpath('//*[@id="container"]/div/div[1]/div/section/div[2]/div[2]/div/p[2]/span/text()').get() or '-'
        
        limit = response.xpath('//*[@id="form1"]/div[1]/div[1]/div/p/text()').get()
        
        description = response.xpath('//div[@class="detail_info"]/p/text()').getall()
        if not description:
            description = response.xpath('//div[@class="detail_info"]/text()').getall()

        description = '\n'.join([elem.strip() for elem in description if elem.strip()])
        
        item_thumbs = response.xpath('//div[@class="item_thumbs_inner"]/ul/li/span/img/@src').getall()

        item = AnimateItem()
        item['datetime'] = datetime.now()
        item['name'] = name 
        item['point'] = point
        item['sku'] = sku
        item['price'] = price
        item['url'] = response.url
        item['description'] = description
        item['sale_start'] = release_date
        item['sale_status'] = stock
        item['bonus'] = bonus
        item['purchase_limit'] = limit
        item['site'] = 'animate'
        item['genre'] = genre
        item['image_urls'] = item_thumbs
        yield item