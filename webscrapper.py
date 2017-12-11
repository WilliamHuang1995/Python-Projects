import bs4, requests, pyperclip

def getAmazonPrice(productUrl):
    res = requests.get(productUrl)
    res.raise_for_status()

    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    elems = soup.select('#priceblock_ourprice')
    elems2 = soup.select('#productTitle')
    return [elems[0].text.strip(),elems2[0].text.strip()]
    
    
url = pyperclip.paste()
price,productName = getAmazonPrice(url)
print('The price for '+productName+' is '+price)
