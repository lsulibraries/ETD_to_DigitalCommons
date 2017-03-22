#! /usr/bin/env python3

import urllib.request
import webbrowser

from bs4 import BeautifulSoup


def find_missing_pdfs(url):
    try:
        html_doc = urllib.request.urlopen(url).read()
    except urllib.error.URLError:
        print('Error:  Unable to connect to the DigitalCommons site')
        quit()
    soup = BeautifulSoup(html_doc, 'html.parser')
    print(soup)

    return [article.find_next('a', href=True)['href'] for article in soup.find_all('p')
            if 'class' in article.attrs and
            'article-listing' in article['class'] and
            "pdf" not in article.find_previous_sibling("p")['class']]


def open_webpages(list_of_urls):
    for url in list_of_urls:
        webbrowser.get(using='google-chrome').open_new_tab(url)


def main_loop(max_page, digcom_coll_url):
    all_urls = ['{}/index.html'.format(digcom_coll_url), ]
    all_urls.append('{}/index.{}.html'.format(digcom_coll_url, i)
                    for i in range(2, max_page))
    all_missing_items = [find_missing_pdfs(url) for url in all_urls]
    flat_missing_items = [item for sublist in all_missing_items for item in sublist]

    print("There were {} items missing a pdf binary".format(len(flat_missing_items)))
    print('These items will now open in your Chrome webbrowser for you to manually attach the pdf.')

    open_webpages(flat_missing_items)


if __name__ == '__main__':
    max_page = 2
    digitalcommons_collection_url = 'http://digitalcommons.lsu.edu/gradschool_disstheses'
    main_loop(max_page, digitalcommons_collection_url)
