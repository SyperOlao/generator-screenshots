import asyncio
from time import sleep
from pyppeteer import launch

url = "https://vk.com/styd.pozor?w=wall-71729358_13128057"
# url = "https://vk.com/instasamka?w=wall-182863061_923576"
# url = "https://vk.com/instasamka?w=wall-182863061_925645"
# url = "https://vk.com/al_feed.php?w=wall-90802042_128826"
# url = "https://vk.com/istoriyadonbassa?w=wall-189776399_233201 "
# url = "https://vk.com/public179350392?w=wall-179350392_9"
# url = "https://vk.com/istoriyadonbassa?w=wall-189776399_233200"
# url = "https://vk.com/delphine_chik?w=wall515103149_134"
# url = "https://vk.com/@-194707757-rss-1610825261-1959186610#"
# url = "https://vk.com/wall-141359201_1489951"
url = "https://vk.com/wall-93250065_1198328"


class GenerateScreenshot:
    pass


async def delete_element_by_id(id: str, page):
    await page.evaluate(f'document.getElementById("{id}")?.remove();')


async def delete_element_by_class_name(class_name: str, page):
    await page.evaluate(
        f"""() => {{
        var elements = document.getElementsByClassName("{class_name}");
        if(!elements) return;
        while(elements?.length >  0){{
            elements[0]?.parentNode?.removeChild(elements[0]);
        }}
    }}"""
    )


async def delete_by_class_names(class_names, post):
    for class_name in class_names:
        await delete_element_by_class_name(class_name, post)


async def delete_by_ids(ids, post):
    for id in ids:
        await delete_element_by_id(id, post)


async def check_existing_element_by_id(id: str, page):
    element = await page.querySelector(
        f"{id}",
    )
    return element is not None


async def get_existing_element(ids, page):
    for id in ids:
        if await check_existing_element_by_id(id, page):
            return id
    return None


# wl_replies_block_wrap
async def main():
    browser = await launch()
    page = await browser.newPage()

    await page.goto(url, {"waitUntil": "load", "timeout": 30000})  # Remove the timeout

    id = await get_existing_element(["#wk_content", "#wide_column"], page)
    print(id)

    if id is None:
        print("Id is not exists")
        return

    try:
        # Try to wait for the selector
        await page.waitForSelector(id)
    except TimeoutError:
        # Handle the timeout exception
        print(f"Timeout while waiting for '{id}'")
        raise

    await delete_by_class_names(
        [
            "wl_replies_block_wrap",
            "wl_replies_wrap",
            "wl_replies",
            "replies",
            "post_replies_header",
        ],
        page,
    )

    await delete_by_ids(["wl_replies_wrap", "page_bottom_banners_root",], page)

    element = await page.querySelector(
        id,
    )

    if element:
        bounding_box = await element.boundingBox()
        print(bounding_box)

    padding = 24
    clip = {
        "width": int(bounding_box["width"]),
        "height": int(bounding_box["height"] + padding),
    }
    await page.setViewport(clip)

    clip_1 = {
        "x": 0,
        "y": padding,
        "width": int(bounding_box["width"]),
        "height": int(bounding_box["height"]),
    }

    if id == "#wide_column":
        
        clip = {
            "width": 1280,
           "height": int(bounding_box["height"] + padding),
        }
        await page.setViewport(clip)
        side_bar = await get_existing_element(["#side_bar"], page)
        x =  int(bounding_box["x"])
    

        if side_bar is not None:
            side_bar_element = await page.querySelector(
                side_bar,
            )
            side_bar_bounding_box = await side_bar_element.boundingBox()
            x += int(side_bar_bounding_box["width"])
        clip_1 = {
            "x": x,
            "y": int(bounding_box["y"]),
            "width": int(bounding_box["width"]),
            "height": int(bounding_box["height"]),
        }


    sleep(1)
    await page.screenshot({"path": "test.png", "clip": clip_1})
    await browser.close()


asyncio.run(main())


# wl_replies_wrap
