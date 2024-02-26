from generate_screenshot.base_screensot import GenerateScreenshot
from config.config import logger

from time import sleep
from pyppeteer.errors import NetworkError


class GenerateScreenshotVk(GenerateScreenshot):
    padding = 24

    async def login_vk(self, email, password):
        try:
            page = await self.browser.newPage()
            await page.goto('https://vk.com/')

            await page.click('#index_email')

            await page.type('#index_email', email)
            await page.click('button[type="submit"]')
            sleep(1)
            name_password = '[name="password"]'
            await page.waitForSelector(name_password)
            await page.type(name_password, password)
            sleep(1)
            await page.click('button[type="submit"]')
            await page.waitForNavigation()
        except NetworkError as e:
            logger.error(e)
            await self.browser_close()

    async def generate_screen_shot(self, url: str, screen_shot_path: str):
        pages = await self.browser.pages()
        for page in pages:
            await page.close({'runBeforeUnload': True})
        self.page = await self.browser.newPage()
        await self.page.goto(url, {"waitUntil": "load", "timeout": 20000, "networkidle0": True})

        id = await self._get_existing_element(
            ["#wk_content", "#wide_column", ".article_layer__views"]
        )
        print(id)

        if id is None:
            raise ValueError(f"Id is not exists")

        try:
            await self.page.waitForSelector(id)
        except TimeoutError:
            raise ValueError(f"Timeout while waiting for '{id}'")

        await self._delete_by_class_names(
            [
                "wl_replies_block_wrap",
                "wl_replies_wrap",
                "wl_reply_form_wrap",
                "wl_replies",
                "replies",
                "post_replies_header",
                "wl_replies_block_wrap"
            ]
        )

        await self._delete_by_ids(
            [
                "wl_replies_wrap",
                "page_bottom_banners_root",
            ]
        )

        element = await self.page.querySelector(
            id,
        )

        bounding_box = await element.boundingBox()
        print(bounding_box)

        clip_for_screen_shot = await self._post_is_wk_content(bounding_box)

        if id == "#wide_column":
            clip_for_screen_shot = await self._post_is_wall(bounding_box)

        sleep(1)
        await self.page.screenshot({"path": screen_shot_path, "clip": clip_for_screen_shot})

    async def _post_is_wk_content(self, bounding_box):
        clip = {
            "width": int(bounding_box["width"]),
            "height": int(bounding_box["height"] + self.padding),
        }
        await self.page.setViewport(clip)

        return {
            "x": 0,
            "y": self.padding,
            "width": int(bounding_box["width"]),
            "height": int(bounding_box["height"]),
        }

    async def _post_is_wall(self, bounding_box):
        clip = {
            "width": 1280,
            "height": int(bounding_box["height"] + self.padding),
        }
        await self.page.setViewport(clip)
        side_bar = await self._get_existing_element(["#side_bar"])
        x = int(bounding_box["x"])

        if side_bar is not None:
            side_bar_element = await self.page.querySelector(
                side_bar,
            )
            side_bar_bounding_box = await side_bar_element.boundingBox()
            x += int(side_bar_bounding_box["width"])
        return {
            "x": x,
            "y": int(bounding_box["y"]),
            "width": int(bounding_box["width"]),
            "height": int(bounding_box["height"]),
        }
