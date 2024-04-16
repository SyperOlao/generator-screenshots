from time import sleep
from pyppeteer import launch
from generate_screenshot.base_screenshot import GenerateScreenshot
import asyncio


class GenerateScreenshotOk(GenerateScreenshot):
    padding = 100

    async def generate_screen_shot(self, url: str, screen_shot_path: str):
        self.browser = await launch()
        self.page = await self.browser.newPage()

        await self.page.goto(url, {"waitUntil": "load", "timeout": 30000})

        self.page.on("dialog", lambda dialog: asyncio.ensure_future(dialog.accept()))

        id = await self._get_existing_element(
            [
                ".media-layer_hld",
            ]
        )
        print(id)

        if id is None:
            await self.browser.close()
            raise ValueError(f"Id is not exists")

        try:
            await self.page.waitForSelector(id)
        except TimeoutError:
            await self.browser.close()
            raise ValueError(f"Timeout while waiting for '{id}'")

        await self._delete_by_class_names(
            [
                "mlr_disc",
                "auth-login-banner__e19f0"
            ]
        )

        element = await self.page.querySelector(
            id,
        )

        bounding_box = await element.boundingBox()
        print(bounding_box)

        clip_for_screen_shot = await self._post_is_tg_content(bounding_box)

        sleep(1)
        await self.page.screenshot(
            {"path": screen_shot_path, "clip": clip_for_screen_shot}
        )
        await self.browser.close()

    async def _post_is_tg_content(self, bounding_box):
        clip = {
            "width": int(bounding_box["width"]),
            "height": int(bounding_box["height"] + self.padding),
        }
        await self.page.setViewport(clip)

        return {
            "x": 0,
            "y": int(bounding_box["y"]),
            "width": int(bounding_box["width"]),
            "height": int(bounding_box["height"]),
        }
