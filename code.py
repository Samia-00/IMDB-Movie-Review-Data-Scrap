import requests,openpyxl  
from bs4 import BeautifulSoup

excel= openpyxl.Workbook()
print(excel.sheetnames)
sheet = excel.active
sheet.title = 'Top Rated Movies'
print(excel.sheetnames)
sheet.append(['Movie Name','Rank','Year of Release','Rating'])

try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()

    soup = BeautifulSoup(source.text,'html.parser')
    #print(soup)

    movies = soup.find('tbody',class_="lister-list").find_all('tr')
    print(len(movies))
    for movie in movies:
        name = movie.find('td',class_ ="titleColumn").a.text 
        #rank = movie.find('td',class_="titleColumn").get_text(strip=True)
        rank= movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0]
        year= movie.find('td',class_="titleColumn").span.text.strip('()')
        rating= movie.find('td',class_='ratingColumn imdbRating').strong.text
        print(name,rank,year,rating)
        sheet.append([name,rank,year,rating])
        #break
except Exception as e:
    print(e)

excel.save('IMDB Movie rating.xlxs')