from CustomSelenium import CustomSelenium
from SitePage import SitePage
from robocorp.tasks import task
from robocorp import workitems


@task
def main_task():
    for item in workitems.inputs:
        website = item.payload['website']
        month_filter = item.payload['month_filter']
        search_phrase = item.payload['search_phrase']
    selenium = CustomSelenium()
    selenium.set_webdriver()
    selenium.open_url(website)
    site_page = SitePage(selenium.driver)
    site_page.search(search_phrase)
    site_page.get_articles_info(int(month_filter), search_phrase)
    print("Done.")
