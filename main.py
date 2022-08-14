import statistics
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import openpyxl
import pathlib

class PythonTest:

    def __init__(self):

        # Starting Process
        print(f'Starting Process...')

        # Driver Settings
        self.userprofile = os.environ['USERPROFILE']
        self.driver = webdriver.Chrome(executable_path=fr"{self.userprofile}\Downloads\chromedriver.exe")
        self.driver.maximize_window()
        url = r'https://twitter.com/explore'
        self.driver.get(url)


    def main(self):

        # Create Excel File
        self.create_excel_file()

        # Define XPath
        self.define_xpath()

        # List of Users to Search
        twitter_users = ["neymarjr", "Twitter", "Anitta", "whindersson", "NetflixBrasil", "BruMarquezine", "maisa",
                         "anaclaraac", "gioewbank", "taisdeverdade", "ANAMARIABRAGA", "gio_antonelli",
                         "julianapaes", "ZAMENZA", "Cristiano", "realmadrid", "Benzema", "ToniKroos", "FIFAWorldCup", "premierleague",
                         "ChampionsLeague", "elonmusk", "BillGates", "JeffBezos", "NASA", "SpaceX", "Tesla",
                         "KingJames", "StephenCurry30", "NBA", "KDTrey5", "Giannis_An34", "KlayThompson",
                         "spidadmitchell", "FCHWPO", "jaytatum0", "MichaelPhelps", "usainbolt", "Ibra_official",
                         "FCBarcelona", "ErlingHaaland", "KMbappe", "NBAHistory", "NBATV", "WashWizards", "BrooklynNets",
                         "Lakers", "JaMorant", "ZO2_", "AntDavis23", "SHAQ", "DwightHoward", "russwest44",
                         "business", "WSJ", "nytimes", "CNN", "cnnbrk", "BBCBreaking", "TheEconomist",
                         "BBCWorld", "SportsCenter", "espn", "ESPNFC", "ESPNPlus", "TomBrady"]
        for user in twitter_users:
            print("User: ", user)
            self.get_user_data(user)
            self.get_user_tweets()

            # Increment counter
            self.last_empty_row += 1

        # Ending process
        self.close_driver()
        print("End Process")


    def get_user_data(self, user_twitter):
        url_user = fr'https://twitter.com/{user_twitter}'
        self.driver.get(url_user)
        time.sleep(6)
        delay = 3  # seconds
        try:
            load_page = WebDriverWait(self.driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, self.css_displayed_name)))
            print("Page is ready")
        except TimeoutException:
            print("Loading took too much time")
        print("Search User: ", user_twitter)

        # displayed_name
        displayed_name = self.driver.find_element(By.CSS_SELECTOR, self.css_displayed_name)
        displayed_name = displayed_name.text.strip()
        print("displayed_name: ", displayed_name)

        # description
        try:
            description = self.driver.find_element(By.CSS_SELECTOR, self.css_description)
            description = description.text.strip()
            print("description: ", description)
        except Exception as e:
            description = '-'
            print("Description no exists - ", e)

        # number_following
        number_following = self.driver.find_element(By.CSS_SELECTOR, self.css_following)
        number_following = number_following.text.splitlines()[0].replace('Seguindo', '').strip()
        print("number_following: ", number_following)

        # number_followers
        number_followers = self.driver.find_element(By.CSS_SELECTOR, self.css_followers)
        number_followers = number_followers.text.splitlines()[1].replace('Seguidores', '').strip()
        print("number_following: ", number_followers)

        # birthday
        try:
            birthday = self.driver.find_element(By.CSS_SELECTOR, self.css_birthday)
            birthday = birthday.text.strip()
            print("birthday: ", birthday)
        except Exception as e:
            birthday = '-'
            print("Birthday no exists - ", e)

        # date_joined
        try:
            date_joined = self.driver.find_element(By.CSS_SELECTOR, self.css_date_joined)
            date_joined = date_joined.text.strip()
            print("date_joined: ", date_joined)
        except Exception as e:
            date_joined = '-'
            print("Date Joined Twitter no exists - ", e)

        # website
        try:
            website = self.driver.find_element(By.CSS_SELECTOR, self.css_website)
            website = website.text.strip()
            print("website: ", website)
        except Exception as e:
            website = '-'
            print("Website no exists - ", e)

        self.write_user_excel(user_twitter, displayed_name, description, number_following, number_followers, birthday, date_joined, website)


    def get_user_tweets(self):
        print('Getting user tweets...')

        try:
            tweet_data = ''
            # self.driver.execute_script("window.scrollTo(0, window.scrollY + 5000)")
            self.driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            time.sleep(2)
            tweets_list = self.driver.find_elements(By.CSS_SELECTOR, "div[role='group']")
            tweet_quantity = len(tweets_list)
            print(tweet_quantity)
            tweet_sum_favorites = 0
            tweet_sum_retweets = 0
            tweet_sum_replies = 0
            tweet_list_favorites = []
            tweet_list_retweets = []
            tweet_list_replies = []
            for tweet in tweets_list:
                tweet_data = tweet_data + '\n' + tweet.text
                tweet_numbers = tweet.accessible_name.upper().replace('RESPOSTAS', '').replace('RETWEETS', '').replace('CURTIDAS', '').strip()

                # get values
                tweet_favorites = float(tweet_numbers.split(',')[0].strip())
                tweet_retweets = float(tweet_numbers.split(',')[1].strip())
                tweet_replies = float(tweet_numbers.split(',')[2].strip())

                # get sum values
                tweet_sum_favorites = tweet_sum_favorites + tweet_favorites
                tweet_sum_retweets = tweet_sum_retweets + tweet_retweets
                tweet_sum_replies = tweet_sum_replies + tweet_replies

                # get median list
                tweet_list_favorites.append(tweet_favorites)
                tweet_list_retweets.append(tweet_retweets)
                tweet_list_replies.append(tweet_replies)

            # get final mean data
            tweet_mean_favorites = tweet_sum_favorites / tweet_quantity
            print(tweet_mean_favorites)
            tweet_mean_retweets = tweet_sum_retweets / tweet_quantity
            print(tweet_mean_retweets)
            tweet_mean_replies = tweet_sum_replies / tweet_quantity
            print(tweet_mean_replies)

            # get final median data
            tweet_median_favorites = statistics.median(tweet_list_favorites)
            tweet_median_retweets = statistics.median(tweet_list_retweets)
            tweet_median_replies = statistics.median(tweet_list_replies)

            self.write_tweets_excel(tweet_data, tweet_sum_favorites, tweet_sum_retweets, tweet_sum_replies, tweet_mean_favorites, tweet_mean_retweets, tweet_mean_replies, tweet_median_favorites, tweet_median_retweets, tweet_median_replies)
        except Exception as e:
            print("Get tweet data failed - ", e)


    def create_excel_file(self):
        print('Creating Excel...')
        self.path = pathlib.Path().resolve()
        print(self.path)
        self.excel_file = openpyxl.Workbook()
        self.sheet = self.excel_file.active
        self.last_empty_row = int(2)
        print('last_empty_row: ', self.last_empty_row)
        self.sheet['A1'].value = 'Username'
        self.sheet['B1'].value = 'Displayed Name'
        self.sheet['C1'].value = 'Description'
        self.sheet['D1'].value = 'Number Of Followers'
        self.sheet['E1'].value = 'Number Of Following'
        self.sheet['F1'].value = 'Birthday'
        self.sheet['G1'].value = 'Data Joined Twitter'
        self.sheet['H1'].value = 'Website'
        self.sheet['I1'].value = 'Last Tweets'
        self.sheet['J1'].value = 'Sum of Favorites'
        self.sheet['K1'].value = 'Sum of Retweets'
        self.sheet['L1'].value = 'Sum of Replies'
        self.sheet['M1'].value = 'Mean of Favorites'
        self.sheet['N1'].value = 'Mean of Retweets'
        self.sheet['O1'].value = 'Mean of Replies'
        self.sheet['P1'].value = 'Median of Favorites'
        self.sheet['Q1'].value = 'Median of Retweets'
        self.sheet['R1'].value = 'Median of Replies'
        time.sleep(1)
        self.excel_file.save(f'{self.path}' + '/Twitter_Report.xlsx')


    def define_xpath(self):
        # Define XPath
        self.css_displayed_name = "div[data-testid='UserName']"
        self.css_description = "div[data-testid='UserDescription']"
        self.css_following = "div[class='css-1dbjc4n r-13awgt0 r-18u37iz r-1w6e6rj']"
        self.css_followers = "div[class='css-1dbjc4n r-13awgt0 r-18u37iz r-1w6e6rj']"
        self.css_birthday = "span[data-testid='UserBirthdate']"
        self.css_date_joined = "span[data-testid='UserJoinDate']"
        self.css_website = "a[data-testid='UserUrl']"


    def write_user_excel(self, user_twitter, displayed_name, description, number_following, number_followers, birthday, date_joined, website):
        print('Writing Excel User Data...')
        self.sheet['A' + str(self.last_empty_row)].value = str(user_twitter)
        self.sheet['B' + str(self.last_empty_row)].value = str(displayed_name)
        self.sheet['C' + str(self.last_empty_row)].value = str(description)
        self.sheet['D' + str(self.last_empty_row)].value = str(number_followers)
        self.sheet['E' + str(self.last_empty_row)].value = str(number_following)
        self.sheet['F' + str(self.last_empty_row)].value = str(birthday)
        self.sheet['G' + str(self.last_empty_row)].value = str(date_joined)
        self.sheet['H' + str(self.last_empty_row)].value = str(website)
        time.sleep(1)
        try:
            self.excel_file.save(f'{self.path}' + '/Twitter_Report.xlsx')
            print("Sucess to save excel file")
        except Exception as e:
            print("Something wrong to save excel file: ", e)


    def write_tweets_excel(self, tweets_text, tweet_sum_favorites, tweet_sum_retweets, tweet_sum_replies, tweet_mean_favorites, tweet_mean_retweets, tweet_mean_replies, tweet_median_favorites, tweet_median_retweets, tweet_median_replies):
        print('Writing Excel Tweets Data...')
        self.sheet['I' + str(self.last_empty_row)].value = tweets_text
        self.sheet['J' + str(self.last_empty_row)].value = tweet_sum_favorites
        self.sheet['K' + str(self.last_empty_row)].value = tweet_sum_retweets
        self.sheet['L' + str(self.last_empty_row)].value = tweet_sum_replies
        self.sheet['M' + str(self.last_empty_row)].value = tweet_mean_favorites
        self.sheet['N' + str(self.last_empty_row)].value = tweet_mean_retweets
        self.sheet['O' + str(self.last_empty_row)].value = tweet_mean_replies
        self.sheet['P' + str(self.last_empty_row)].value = tweet_median_favorites
        self.sheet['Q' + str(self.last_empty_row)].value = tweet_median_retweets
        self.sheet['R' + str(self.last_empty_row)].value = tweet_median_replies
        try:
            self.excel_file.save(f'{self.path}' + '/Twitter_Report.xlsx')
            print("Sucess to save excel file")
        except Exception as e:
            print("Something wrong to save excel file: ", e)


    def close_driver(self):
        self.driver.close()
        self.driver.quit()


if __name__ == '__main__':
    PythonTest().main()