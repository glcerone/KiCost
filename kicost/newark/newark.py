# MIT license
#
# Copyright (C) 2015 by XESS Corporation
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

# Inserted by Pasteurize tool.
from __future__ import print_function
from __future__ import unicode_literals
from __future__ import division
from __future__ import absolute_import
from builtins import zip
from builtins import range
from builtins import int
from builtins import str
from future import standard_library
standard_library.install_aliases()

import future

import sys
import pprint
import re
import difflib
import logging
import tqdm
from bs4 import BeautifulSoup
from random import randint
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range, xl_range_abs
from yattag import Doc, indent  # For generating HTML page for local parts.
import multiprocessing
from multiprocessing import Pool # For running web scrapes in parallel.
import http.client # For web scraping exceptions.

try:
    from urllib.parse import urlencode, quote as urlquote, urlsplit, urlunsplit
    import urllib.request
    from urllib.request import urlopen, Request
except ImportError:
    from urlparse import quote as urlquote, urlsplit, urlunsplit
    from urllib import urlencode
    from urllib2 import urlopen, Request

from ..kicost import PartHtmlError, FakeBrowser
from ..kicost import logger, DEBUG_OVERVIEW, DEBUG_DETAILED, DEBUG_OBSESSIVE

from currency_converter import CurrencyConverter

SEPRTR = ':'  # Delimiter between library:component, distributor:field, etc.

HTML_RESPONSE_RETRIES = 2 # Num of retries for getting part data web page.

WEB_SCRAPE_EXCEPTIONS = (urllib.request.URLError, http.client.HTTPException)

from ..kicost import distributors
distributors.update(
    {
        'newark': {
            'scrape': 'web',
            'function': 'newark',
            'label': 'Newark',
            'order_cols': ['part_num', 'purch', 'refs'],
            'order_delimiter': ','
        }
    }
)


def get_newark_price_tiers(html_tree):
    '''Get the pricing tiers from the parsed tree of the Newark product page.'''
    price_tiers = {}
    try:
        qty_strs = []
        for qty in html_tree.find(
            'table',
            class_=('tableProductDetailPrice', 'pricing')).find_all(
                'td',
                class_='qty'):
            qty_strs.append(qty.text)
        price_strs = []
        for price in html_tree.find(
            'table',
            class_=('tableProductDetailPrice', 'pricing')).find_all(
                'td',
                class_='threeColTd'):
            price_strs.append(price.text)
        qtys_prices = list(zip(qty_strs, price_strs))
        for qty_str, price_str in qtys_prices:
            try:
                qty = re.search('(\s*)([0-9,]+)', qty_str).group(2)
                qty = int(re.sub('[^0-9]', '', qty))
                price_tiers[qty] = float(re.sub('[^0-9\.]', '', price_str))
            except (TypeError, AttributeError, ValueError):
                continue
    except AttributeError:
        # This happens when no pricing info is found in the tree.
        logger.log(DEBUG_OBSESSIVE, 'No Newark pricing information found!')
        return price_tiers  # Return empty price tiers.
    return price_tiers


def get_newark_part_num(html_tree):
    '''Get the part number from the Newark product page.'''
    try:
        # Newark catalog number is stored in a description list, so get
        # all the list terms and descriptions, strip all the spaces from those,
        # and pair them up.
        div = html_tree.find('div', class_='productDescription').find('dl')
        dt = [re.sub('\s','',d.text) for d in div.find_all('dt')]
        dd = [re.sub('\s','',d.text) for d in div.find_all('dd')]
        dtdd = {k:v for k,v in zip(dt,dd)}  # Pair terms with descriptions.
        return dtdd['NewarkPartNo.:']
    except KeyError:
        logger.log(DEBUG_OBSESSIVE, 'No Newark catalog number found!')
        return '' # No catalog number found in page.
    except AttributeError:
        logger.log(DEBUG_OBSESSIVE, 'No Newark product description found!')
        return '' # No ProductDescription found in page.


def get_newark_qty_avail(html_tree):
    '''Get the available quantity of the part from the Newark product page.'''
    try:
        # Note that 'availability' is misspelled in the container class name!
        qty_str = html_tree.find('div',
                                 class_='avalabilityContainer').find(
                                     'span',
                                     class_='availability').text
    except (AttributeError, ValueError):
        # No quantity found (not even 0) so this is probably a non-stocked part.
        # Return None so the part won't show in the spreadsheet for this dist.
        return None
    try:
        qty = re.sub('[^0-9]','',qty_str)  # Strip all non-number chars.
        return int(re.sub('[^0-9]', '', qty_str))  # Return integer for quantity.
    except ValueError:
        # No quantity found (not even 0) so this is probably a non-stocked part.
        # Return None so the part won't show in the spreadsheet for this dist.
        logger.log(DEBUG_OBSESSIVE, 'No Newark part quantity found!')
        return None


def get_newark_part_html_tree(dist, pn, extra_search_terms='', url=None, descend=2, local_part_html=None):
    '''Find the Newark HTML page for a part number and return the URL and parse tree.'''

    # Use the part number to lookup the part using the site search function, unless a starting url was given.
    if url is None:
        url = 'http://www.newark.com/webapp/wcs/stores/servlet/Search?catalogId=15003&langId=-1&storeId=10194&gs=true&st=' + urlquote(
            pn + ' ' + extra_search_terms,
            safe='')
    elif url[0] == '/':
        url = 'http://www.newark.com' + url
    elif url.startswith('..'):
        url = 'http://www.newark.com/Search/' + url

    # Open the URL, read the HTML from it, and parse it into a tree structure.
    for _ in range(HTML_RESPONSE_RETRIES):
        try:
            req = FakeBrowser(url)
            response = urlopen(req)
            html = response.read()
            break
        except WEB_SCRAPE_EXCEPTIONS:
            logger.log(DEBUG_DETAILED,'Exception while web-scraping {} from {}'.format(pn, dist))
            pass
    else: # Couldn't get a good read from the website.
        logger.log(DEBUG_OBSESSIVE,'No HTML page for {} from {}'.format(pn, dist))
        raise PartHtmlError

    try:
        tree = BeautifulSoup(html, 'lxml')
    except Exception:
        logger.log(DEBUG_OBSESSIVE,'No HTML tree for {} from {}'.format(pn, dist))
        raise PartHtmlError

    # Abort if the part number isn't in the HTML somewhere.
    # (Only use the numbers and letters to compare PN to HTML.)
    if re.sub('[\W_]','',str.lower(pn)) not in re.sub('[\W_]','',str.lower(str(html))):
        logger.log(DEBUG_OBSESSIVE,'No part number {} in HTML page from {}'.format(pn, dist))
        raise PartHtmlError

    # If the tree contains the tag for a product page, then just return it.
    if tree.find('div', class_='productDisplay', id='page') is not None:
        return tree, url

    # If the tree is for a list of products, then examine the links to try to find the part number.
    if tree.find('table', class_='productLister', id='sProdList') is not None:
        logger.log(DEBUG_OBSESSIVE,'Found product table for {} from {}'.format(pn, dist))
        if descend <= 0:
            logger.log(DEBUG_OBSESSIVE,'Passed descent limit for {} from {}'.format(pn, dist))
            raise PartHtmlError
        else:
            # Look for the table of products.
            products = tree.find('table',
                                 class_='productLister',
                                 id='sProdList').find('tbody').find_all('tr')

            # Extract the product links for the part numbers from the table.
            product_links = []
            for p in products:
                try:
                    product_links.append(
                        p.find('td', class_='mftrPart').find('a'))
                except AttributeError:
                    continue

            # Extract all the part numbers from the text portion of the links.
            part_numbers = [l.text for l in product_links]

            # Look for the part number in the list that most closely matches the requested part number.
            try:
                match = difflib.get_close_matches(pn, part_numbers, 1, 0.0)[0]
            except IndexError:
                raise PartHtmlError

            # Now look for the link that goes with the closest matching part number.
            for l in product_links:
                if l.text == match:
                    # Get the tree for the linked-to page and return that.
                    logger.log(DEBUG_OBSESSIVE,'Selecting {} from product table for {} from {}'.format(l.text, pn, dist))
                    return get_newark_part_html_tree(dist, pn, extra_search_terms,
                                url=l['href'], descend=descend-1)

    # I don't know what happened here, so give up.
    logger.log(DEBUG_OBSESSIVE,'Unknown error for {} from {}'.format(pn, dist))
    raise PartHtmlError
