from ipaddress import NetmaskValueError
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import requests
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import bs4 as bs
from datetime import date
from datetime import timedelta
from dateutil.relativedelta import relativedelta
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.stem.wordnet import WordNetLemmatizer
from spellchecker import SpellChecker
import string
from collections import Counter
Reviews = pd.read_excel('C:\\Users\\George.Burns\\Documents\\Projects\\WWT Customer Reviews Data.xlsx', sheet_name='TripAdvisor')
Google = pd.read_excel('C:\\Users\\George.Burns\\Documents\\Projects\\WWT Customer Reviews Data.xlsx', sheet_name='Google Reviews')
url = "https://www.tripadvisor.co.uk/Attraction_Review-g1902845-d3597529-Reviews-WWT_Slimbridge_Wetland_Centre-Slimbridge_Cotswolds_England.html"
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
path = "\\Users\\George.Burns\\Downloads\\Python Scripts\\chromedriver.exe"
driver = webdriver.Chrome(path, options=options)
driver.get(url)
time.sleep(4)
driver.find_element(By.XPATH, '//*[@id="onetrust-accept-btn-handler"]').click()
#change to decide the number of pages/scrolls scraped for each respective
pages_to_scrape = 1
scrolls = 1
#Trip Advisor
pages = 1
number = 0
sites = 0
headers = {}
Columns = ('Date of Review', 'Rating', 'Review', 'Website','Site')
reviews = []
ratings = []
dates = []
site = []
website = []
def scraper(url, pages, number, head, tail, name):
    while pages < pages_to_scrape + 1:
        source = requests.get(url, headers=headers)
        html = source.text
        soup = bs.BeautifulSoup(html, 'html.parser')
        for review_selector in soup.find_all('div', class_ = "fIrGe _T bgMZj"):
            review = review_selector.find('span', class_ ="yCeTE")
            if review != None:
                reviews.append(review.text)
                site.append(name)
                website.append('TripAdvisor')
        for review_selector in soup.find_all('div', class_ = '_c'):
            rating = review_selector.find('svg', class_ ="UctUV d H0")
            if rating != None:
                ratings.append(rating['aria-label'])
            date = review_selector.find('div', class_ ="RpeCd")
            if date != None:
                dates.append(date.text)
        if pages < pages_to_scrape: 
            time.sleep(5)
            test = driver.find_elements(By.CLASS_NAME, "xkSty")
            if test == 1:
                driver.find_element(By.CLASS_NAME, "xkSty").click()
        pages += 1
        number += 10
        url = head + str(number) + tail 
while sites < 10:
    if sites == 0:
        name = 'Slimbridge'
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g1902845-d3597529-Reviews-or' 
        tail = '-WWT_Slimbridge_Wetland_Centre-Slimbridge_Cotswolds_England.html'
        print('Slimbridge')
        scraper(url, pages, number, head, tail, name)
    pages = 1
    number = 0
    sites += 1
    if sites == 1:
        name = 'London'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g186338-d187534-Reviews-WWT_London_Wetland_Centre-London_England.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g186338-d187534-Reviews-or' 
        tail = '-WWT_London_Wetland_Centre-London_England.html'
        print('London')
        scraper(url, pages, number, head, tail, name)    
    pages = 1
    number = 0
    sites += 1
    if sites == 2:
        name = 'Martin Mere'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g644362-d261551-Reviews-WWT_Martin_Mere_Wetland_Centre-Burscough_Ormskirk_Lancashire_England.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g644362-d261551-Reviews-or' 
        tail = '-WWT_Martin_Mere_Wetland_Centre-Burscough_Ormskirk_Lancashire_England.html'
        print('Martin Mere')
        scraper(url, pages, number, head, tail, name)      
    pages = 1
    number = 0
    sites += 1
    if sites == 3:
        name = 'Arundel'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g186405-d297089-Reviews-WWT_Arundel_Wetland_Centre-Arundel_Arun_District_West_Sussex_England.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g186405-d297089-Reviews-or' 
        tail = '-WWT_Arundel_Wetland_Centre-Arundel_Arun_District_West_Sussex_England.html'
        print('Arundel')
        scraper(url, pages, number, head, tail, name)   
    pages = 1
    number = 0
    sites += 1
    if sites == 4:
        name = 'Llanelli'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g776264-d219276-Reviews-Llanelli_Wetland_Centre-Llanelli_Carmarthenshire_Wales.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g776264-d219276-Reviews-or'
        tail = '-Llanelli_Wetland_Centre-Llanelli_Carmarthenshire_Wales.html'
        print('Llanelli')
        scraper(url, pages, number, head, tail, name)
    pages = 1
    number = 0
    sites += 1
    if sites == 5:
        name = 'Welney'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g7176430-d1837805-Reviews-WWT_Welney_Wetland_Centre-Welney_Norfolk_East_Anglia_England.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g7176430-d1837805-Reviews-or' 
        tail ='-WWT_Welney_Wetland_Centre-Welney_Norfolk_East_Anglia_England.html'
        print('Welney')
        scraper(url, pages, number, head, tail, name)
    pages = 1
    number = 0
    sites += 1
    if sites == 6:
        name = 'Washington'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g504187-d591659-Reviews-WWT_Washington_Wetland_Centre-Washington_Tyne_and_Wear_England.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g504187-d591659-Reviews-or' 
        tail = '-WWT_Washington_Wetland_Centre-Washington_Tyne_and_Wear_England.html'
        print('Washington')
        scraper(url, pages, number, head, tail, name)    
    pages = 1
    number = 0
    sites += 1
    if sites == 7:
        name = 'Castle Espie'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g551732-d318705-Reviews-WWT_Castle_Espie_Wetland_Centre-Comber_County_Down_Northern_Ireland.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g551732-d318705-Reviews-or' 
        tail = '-WWT_Castle_Espie_Wetland_Centre-Comber_County_Down_Northern_Ireland.html'
        print('Castle Espie')
        scraper(url, pages, number, head, tail, name) 
    pages = 1
    number = 0
    sites += 1
    if sites == 8:
        name = 'Caerlaverock'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g186513-d2038192-Reviews-WWT_Caerlaverock_Wetland_Centre-Dumfries_Dumfries_and_Galloway_Scotland.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g186513-d2038192-Reviews-or' 
        tail = '-WWT_Caerlaverock_Wetland_Centre-Dumfries_Dumfries_and_Galloway_Scotland.html'
        print('Caerlaverock')
        scraper(url, pages, number, head, tail, name) 
    pages = 1
    number = 0
    sites += 1
    if sites == 9:
        name = 'Steart'
        url = "https://www.tripadvisor.co.uk/Attraction_Review-g504115-d10091988-Reviews-WWT_Steart_Marshes-Bridgwater_Somerset_England.html"
        head = 'https://www.tripadvisor.co.uk/Attraction_Review-g504115-d10091988-Reviews-or'
        tail = '-WWT_Steart_Marshes-Bridgwater_Somerset_England.html'
        print('Steart')
        scraper(url, pages, number, head, tail, name)
    pages = 1
    number = 0
    sites += 1
print('TripAdvisor Complete')
Temp = list(zip(dates, ratings, reviews, website, site))
Temp = pd.DataFrame(Temp,columns=Columns)
for review in Temp:
    Temp['Date of Review'] = Temp['Date of Review'].str[:8]
    Temp['Rating'] = Temp['Rating'].str[:1]
Temp['Rating'] = pd.to_numeric(Temp['Rating'])
Reviews = pd.concat([Reviews, Temp])
Reviews = Reviews.drop_duplicates(keep='first')
print(Reviews)
#Google Reviews
reviews = []
ratings = []
dates = []
website = []
site = []
names = []
def Gscraper(url, centre, sites):
    driver.get(url)
    time.sleep(2)
    if sites == 0:
        driver.find_element(By.XPATH,'//*[@id="yDmH0d"]/c-wiz/div/div/div/div[2]/div[1]/div[3]/div[1]/div[1]/form[2]/div/div/button').click()
    time.sleep(2)
    button = driver.find_element(By.CSS_SELECTOR,'[aria-label="Sort reviews"]')
    time.sleep(1)
    ActionChains(driver).move_to_element(button).click(button).perform()
    time.sleep(2)
    options = driver.find_elements(By.CLASS_NAME,'fxNQSd')
    options[1].click()
    time.sleep(2)
    last_height = driver.execute_script("return document.body.scrollHeight")
    number = 0
    while True:
        number = number+1
        ele = driver.find_element(By.XPATH,'//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]')
        driver.execute_script('arguments[0].scrollBy(0, 5000);', ele)
        time.sleep(3)
        print(f'last height: {last_height}')
        ele = driver.find_element(By.XPATH,'//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]')
        new_height = driver.execute_script("return arguments[0].scrollHeight", ele)
        print(f'new height: {new_height}')
        if number == scrolls:
            break
        if new_height == last_height:
            break
        print('cont')
        last_height = new_height
    item = driver.find_elements(By.XPATH, '//*[@id="QA0Szd"]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[10]')
    time.sleep(3)
    for i in item:
        button = i.find_elements(By.TAG_NAME, 'button')
        for m in button:
            if m.text == 'More':
                m.click()
        time.sleep(5)
    page_source = driver.page_source
    soup = bs.BeautifulSoup(page_source, 'html.parser')
    reviews_selector = soup.find_all('div', class_='jJc9Ad')
    for review_selector in reviews_selector:
        name = review_selector.find('div', class_='d4r55').get_text()
        name = name.strip()
        names.append(name)
        review = review_selector.find('span', class_='wiI7pd').get_text()
        review = review.strip()
        reviews.append(review)
        rating = review_selector.find('span', class_='kvMYJc')
        rating = rating['aria-label'].strip()
        ratings.append(rating)
        date = review_selector.find('span', class_='rsqaWe').get_text()
        date = date.strip()
        dates.append(date)
        website.append('Google Reviews')
        site.append(centre)
sites = 0
while sites < 10:
    if sites == 0:
        url = "https://www.google.com/maps/place/WWT+Slimbridge/@51.7400167,-2.4076206,17z/data=!4m7!3m6!1s0x4871a14b73facfed:0xa46272e0a3b0e2ae!8m2!3d51.7400167!4d-2.4054319!9m1!1b1"
        centre = 'Slimbridge'
        Gscraper(url, centre, sites)
        print('Slimbridge')
        sites += 1
    if sites == 1:
        url = "https://www.google.com/maps/place/WWT+London+Wetland+Centre/@51.4765438,-0.2396951,17z/data=!3m1!4b1!4m5!3m4!1s0x48760fac6312fa9d:0x450942811811d9eb!8m2!3d51.4765405!4d-0.2354679"
        centre = 'London'
        Gscraper(url, centre, sites)
        print('London')
        sites += 1
    if sites == 2:
        url = "https://www.google.com/maps/place/WWT+Martin+Mere/@53.6226228,-2.8673408,17z/data=!3m1!4b1!4m5!3m4!1s0x487b15d397865011:0x135d8fd57b295385!8m2!3d53.6226196!4d-2.8651521"
        centre = 'Martin Mere'
        Gscraper(url, centre, sites)
        print('Martin Mere')
        sites += 1
    if sites == 3:
        url = "https://www.google.com/maps/place/WWT+Arundel/@50.8631418,-0.5534051,17z/data=!3m1!4b1!4m5!3m4!1s0x4875ba98f1689e57:0x8b7fb517165b832!8m2!3d50.8631384!4d-0.5512164"
        centre = 'Arundel'
        Gscraper(url, centre, sites)
        print('Arundel')
        sites += 1
    if sites == 4:
        url = "https://www.google.com/maps/place/WWT+Llanelli+Wetland+Centre/@51.6649809,-4.1273899,17z/data=!3m1!4b1!4m5!3m4!1s0x486ef277738339d9:0xdd45df59c2d15923!8m2!3d51.6649776!4d-4.1252012"
        centre = 'Llanelli'
        Gscraper(url, centre, sites)
        print('Llanelli')
        sites += 1
    if sites == 5:
        url = "https://www.google.com/maps/place/WWT+Welney/@52.5269464,0.2765863,17z/data=!3m1!4b1!4m5!3m4!1s0x47d810a5d41a5b95:0x34913f008968e592!8m2!3d52.5269432!4d0.278775"
        centre = 'Welney'
        Gscraper(url, centre, sites)
        print('Welney')
        sites += 1
    if sites == 6:
        url = "https://www.google.com/maps/place/WWT+Washington/@54.8995988,-1.4907153,17z/data=!3m1!4b1!4m5!3m4!1s0x487e64dbfa569cb5:0x14b03c91635b8514!8m2!3d54.8995958!4d-1.4864881"
        centre = 'Washington'
        Gscraper(url, centre, sites)
        print('Washington')
        sites += 1
    if sites == 7:
        url = "https://www.google.com/maps/place/WWT+Castle+Espie/@54.5302893,-5.6983789,17z/data=!3m1!4b1!4m5!3m4!1s0x4861731c68b1f00d:0xc4029980635c8bf5!8m2!3d54.5302862!4d-5.6961902"
        centre = 'Castle Espie'
        Gscraper(url, centre, sites)
        print('Castle Espie')
        sites += 1
    if sites == 8:
        url = "https://www.google.com/maps/place/WWT+Caerlaverock/@54.9761908,-3.4883265,17z/data=!3m1!4b1!4m5!3m4!1s0x487d3345c6a442f9:0xbd520cff392a9016!8m2!3d54.9761878!4d-3.4840993"
        centre = 'Caerlaverock'
        Gscraper(url, centre, sites)
        print('Caerlaverock')
        sites += 1
    if sites == 9:
        url = "https://www.google.com/maps/place/WWT+Steart+Marshes/@51.1918526,-3.0732238,17z/data=!3m1!4b1!4m5!3m4!1s0x486df89b320f7e91:0x161d16d0ae8526e9!8m2!3d51.1918493!4d-3.0710351"
        centre = 'Steart'
        Gscraper(url, centre, sites)
        print('Steart')
        sites += 1
        driver.quit()
print('Google Reviews Complete')
Columns = ('Date of Review', 'Rating', 'Review', 'Website','Site', 'Name')
Temp = list(zip(dates, ratings, reviews, website, site, names))
Temp = pd.DataFrame(Temp,columns=Columns)
def date_changer(df):
    dates = []
    dates_fixed = []
    dates_perfect = []
    change = {'weeks':'week',
             'months':'month',
             'years':'year',
             'days':'day',
             'a ':'1 '}
    for row in df.itertuples():
        dates.append(row[1])
    for string in dates:
        for word, replacement in change.items():
            date_fixed = string.replace(word, replacement).strip()
        dates_fixed.append(date_fixed)
    for dates in dates_fixed:
        if 'hour' not in dates:
            period = int(dates[:2])
            if 'year' in dates:
                date_P = date.today() - relativedelta(years=period)
            if 'month' in dates:
                date_P = date.today() - relativedelta(months=period)
            if 'day' in dates:
                date_P = date.today() - timedelta(days=period)
            if 'hour' in dates:
                date_P = date.today()
        else:
            date_P = date.today()
        dates_perfect.append(date_P)
    df['Date of Review'] = dates_perfect
date_changer(Temp)
for review in Temp:
    Temp['Rating'] = Temp['Rating'].str[:1]
Temp['Rating'] = pd.to_numeric(Temp['Rating'])
Google = pd.concat([Google, Temp])
Google = Google.drop_duplicates(keep='last') 
Google.drop('Name',inplace=True, axis=1)
# analysing words in the reviews
nltk.download('stopwords')
nltk.download('punkt')
lemmatizer = WordNetLemmatizer()
All = pd.concat([Reviews,Google]).reset_index()
All.drop('index',inplace=True, axis=1)
All = All.dropna(axis=0)
def word_analysis(df, site, rating):
    stop_words = set(stopwords.words('english'))
    if rating > 3:
        rating_df = df[(df['Rating'] >= rating) & (df['Site'] == site)]
    if rating <= 3:
        rating_df = df[(df['Rating'] <= rating) & (df['Site'] == site)]
    texts = []
    txts = []
    ltxts = []
    for row in rating_df.itertuples():
        spell = SpellChecker()
        text = row[3]
        text = text.translate(str.maketrans('', '', string.punctuation))
        misspelled = spell.unknown(text)
        for word in misspelled:
            correction = spell.correction(word)
            text.replace(word,correction)
        texts.append(text.lower())
    print(texts)
    for text in texts:
        wt = word_tokenize(text)
        wtlist = []
        for txt in wt:
            if txt not in stop_words:
                wtlist.append(txt)
        txts.append(wtlist)
    print(txts)
    for text in txts:
        lwt = []
        print(text)
        for word in text:
            new_word = lemmatizer.lemmatize(word)
            lwt.append(new_word)
        ltxts.append(new_word)
    print(ltxts)
    occurence = Counter(x for xs in ltxts for x in set(xs))
    temp = pd.DataFrame(list(occurence.items()))
    for row in temp:
        temp['Rating'] = rating
        temp['Site'] = site
    temp.columns = ['Word', 'Occurence', 'Rating', 'Site']
    temp = temp.sort_values('Occurence', ascending=False)
    print(temp)
    return temp
sites = 0
rating_count = 0
Columns = ['Word', 'Occurence', 'Rating', 'Site']
WordAnalysis = pd.DataFrame(columns = Columns)
while sites < 10:
    if sites == 0:
        site = 'Slimbridge'
        print(site)
        if rating_count == 0:
            rating = 3
            temp = word_analysis(All, site, rating)
            rating_count += 1
        if rating_count == 1:
            rating = 4
            temp2 = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([temp2, temp])
        sites += 1
    rating_count = 0
    if sites == 1:
        site = 'London'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
            rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        sites += 1
    rating_count = 0
    if sites == 2:
        site = 'Martin Mere'
        print(site)
        if rating_count == 0:
            rating = 3
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
            rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        sites += 1
    rating_count = 0
    if sites == 3:
        site = 'Arundel'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        sites += 1
    rating_count = 0
    if sites == 4:
        site = 'Llanelli'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        sites += 1
    rating_count = 0
    if sites == 5:
        site = 'Welney'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        sites += 1
    rating_count = 0
    if sites == 6:
        site = 'Washington'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        sites += 1
    rating_count = 0
    if sites == 7:
        site = 'Castle Espie'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        sites += 1
    rating_count = 0
    if sites == 8:
        site = 'Caerlaverock'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        sites += 1
    rating_count = 0
    if sites == 9:
        site = 'Steart'
        print(site)
        if rating_count == 0:
            rating = 3 
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        if rating_count == 1:
            rating = 4
            temp = word_analysis(All, site, rating)
            WordAnalysis = pd.concat([WordAnalysis, temp])
        rating_count += 1
        sites += 1 
print('Review Analysis Complete')
WordAnalysisPositiveSite = WordAnalysis[WordAnalysis['Rating'] == 4]
WordAnalysisNegativeSite = WordAnalysis[WordAnalysis['Rating'] == 3]
#outputting
writer = pd.ExcelWriter('WWT Customer Reviews Data.xlsx')
Reviews.to_excel(writer, sheet_name='TripAdvisor', index=False)
Google.to_excel(writer, sheet_name='Google Reviews', index=False)
WordAnalysisPositiveSite.to_excel(writer, sheet_name='Positive (Word, Site)', index=False)
WordAnalysisNegativeSite.to_excel(writer, sheet_name='Negative (Word, Site)', index=False)
writer.save()
print('Reviews has been successful exported.')
