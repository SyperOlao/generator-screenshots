from pyppeteer.errors import NetworkError

from config.config import logger
import time
from abc import abstractmethod

from pyppeteer import launch


class GenerateScreenshot:
    def __init__(self):
        self.browser = None
        self.page = None

    async def browser_close(self):
        await self.browser.close()

    async def browser_open(self):
        self.browser = await launch(
            headless=True, args=["--no-sandbox", "--disable-setuid-sandbox"]
        )
        # self.browser = await launch(headless=False)
    async def close_page(self):
        await self.page.waitForNavigation({'timeout': 300})
        await self.page.close()
        
    @abstractmethod
    async def generate_screen_shot(self, url: str, screen_shot_path: str):
        pass

    async def generate_screen_shots(self, urls: list[str], screen_shot_path: str):
        self.page = await self.browser.newPage()
        for url in urls:
            # self.browser.on('disconnected', lambda: asyncio.get_event_loop().run_until_complete(self.browser_open()))
            start_time = time.time()
            name = url.split("/")[-1]
            try:
                await self.generate_screen_shot(url, f"{screen_shot_path}/{name}.png")
                # time.sleep(1)
            except Exception as e:
                logger.warn(f"Screenshot {name} does not exists, {e}")
            end_time = time.time()
            logger.error(f"Execution time: {end_time - start_time} sec for {name}")

    async def _delete_element_by_id(self, id: str):
        await self.page.evaluate(f'document.getElementById("{id}")?.remove();')

    async def _delete_element_by_class_name(self, class_name: str):
        await self.page.evaluate(
            f"""() => {{
            var elements = document.getElementsByClassName("{class_name}");
            if(!elements) return;
            while(elements?.length >  0){{
                elements[0]?.parentNode?.removeChild(elements[0]);
            }}
        }}"""
        )

    async def close_old_pages(self):
        pages = await self.browser.pages()
        try:
            for page in pages:
                await page.close({"runBeforeUnload": True})
        except NetworkError as e:
            logger.error(f"NetworkError: {e}")

    async def _delete_by_class_names(self, class_names):
        for class_name in class_names:
            await self._delete_element_by_class_name(class_name)

    async def _delete_by_ids(self, ids):
        for id in ids:
            await self._delete_element_by_id(id)

    async def _check_existing_element(self, id: str):
        element = await self.page.querySelector(
            f"{id}",
        )
        return element is not None

    async def _get_existing_element(self, ids):
        for id in ids:
            if await self._check_existing_element(id):
                return id
        return None
