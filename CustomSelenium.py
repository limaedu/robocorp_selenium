from RPA.core.webdriver import start
from selenium.webdriver.chrome.options import Options
from robocorp import log


class CustomSelenium:

    def __init__(self):
        self.driver = None
        self._setup_logger()

    def _setup_logger(self) -> None:
        log.setup_log(
            log_level="debug",
            output_log_level="info",
            output_stream="stdout",
        )

    def set_options(self) -> Options:
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--disable-popup-blocking')
        options.add_argument("--start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        return options

    def set_webdriver(self, browser: str = "Chrome") -> None:
        options = self.set_options()
        log.info(f"Initializing {browser}.")
        self.driver = start(browser=browser, options=options)
        self.driver.set_window_size(1440, 900)
        log.info(f"{browser} initialized.")

    def open_url(self, url: str) -> None:
        log.info(f"Opening {url}.")
        self.driver.get(url)
