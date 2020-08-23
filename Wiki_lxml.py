import lxml.html
import requests
from Wikipedia.save_excel import insert_word_into_excel


def next_url():
    pass


def get_info_about_products():
    COUNTRY = 'Ukraine'
    URL = f'https://en.wikipedia.org/wiki/{COUNTRY}'
    html = requests.get(URL)
    doc = lxml.html.fromstring(html.content)
    category_product = doc.xpath('//table[@class="infobox geography vcard"]/tbody')[0]
    find_keywords(body=category_product, country=COUNTRY)


def find_keywords(body, country):
    capital = extract_capital(body=body)
    official_languages = extract_official_languages(body=body)
    government = extract_government(body=body)
    legislature = extract_legislature(body=body)
    religion = extract_religion(body=body)
    ethnic_groups = extract_ethnic_groups(body=body)
    area = extract_area(body=body)

    keywords = {'Capital': capital, 'Official_languages': official_languages, 'Government': government,
                'Legislature': legislature, 'Religion': religion, 'Ethnic_groups': ethnic_groups,
                'Area': area}
    print(keywords)
    save_into_excel(keywords=keywords, country=country)


def extract_capital(body):
    capital = body.xpath(
        ".//th[@scope='row'][contains(text(), 'Capital')]/following-sibling::td//a/text()")[0]
    return capital


def extract_official_languages(body):
    official_languages = body.xpath(
        ".//th[@scope='row'][contains(text(), 'Official languages')]"
        "/following-sibling::td//a[not(ancestor::sup)]/text()")
    if not official_languages:
        official_languages = body.xpath(
            ".//th[@scope='row'][contains(text(), 'Official language')]"
            "/following-sibling::td//a[not(ancestor::sup)]/text()")
    return official_languages


def extract_government(body):
    government = body.xpath(
        ".//th[@scope='row']//a[contains(text(),'Government')]/parent::th/following-sibling::td//a"
        "[not(ancestor::sup)]/text()")
    return government


def extract_legislature(body):
    legislature = body.xpath(
        ".//th[@scope='row'][contains(text(), 'Legislature')]/following-sibling::td//a"
        "[not(ancestor::sup)]/text()")[0]
    return legislature


def extract_religion(body):
    religion = body.xpath(
        ".//th[@scope='row'][contains(text(), 'Religion ')]"
        "/following-sibling::td//a[not(ancestor::sup)]/text()")
    return religion


def extract_ethnic_groups(body):
    ethnic_groups = body.xpath(
        ".//a[@href='/wiki/Ethnic_group']/parent::th/following-sibling::td//a[not(ancestor::sup)]/text()")
    return ethnic_groups


def extract_area(body):
    area = body.xpath(
        ".//tr[@class='mergedtoprow']//th/a[contains(text(), 'Area ')]"
        "/parent::th/parent::tr/following-sibling::tr[@class='mergedrow']//td[not(ancestor::sup)]/text()")
    return area[0]


def save_into_excel(keywords, country):
    insert_word_into_excel(category={'type': 'Country', 'country': country}, keywords=keywords)


if __name__ == '__main__':
    get_info_about_products()
