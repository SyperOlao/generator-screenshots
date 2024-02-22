import asyncio
from generate_screenshot.ok.generate_ok import GenerateScreenshotOk
from generate_screenshot.telegram.generate_tg import GenerateScreenshotTg
from generate_screenshot.vk.generate_vk import GenerateScreenshotVk
import time

urls_vk = [
    "https://vk.com/public222510135?w=wall-222510135_41",
    "https://vk.com/styd.pozor?w=wall-71729358_13128057",
    "https://vk.com/instasamka?w=wall-182863061_923576",
    "https://vk.com/instasamka?w=wall-182863061_925645",
    "https://vk.com/al_feed.php?w=wall-90802042_128826",
    "https://vk.com/istoriyadonbassa?w=wall-189776399_233201",
    "https://vk.com/public179350392?w=wall-179350392_9",
    "https://vk.com/istoriyadonbassa?w=wall-189776399_233200",
    "https://vk.com/delphine_chik?w=wall515103149_134",
    "https://vk.com/@-194707757-rss-1610825261-1959186610#",
    "https://vk.com/wall-141359201_1489951",
    "https://vk.com/wall-93250065_1198328",
    "https://vk.com/wall-224359326_12",
]

urls_tg = [
    "https://telegram.me/sakhadaychat/591993",
    # "https://telegram.me/lenskiy_vestni/29702",
]

urls_ok = [
    "https://ok.ru/group/70000004708458/topic/156795335883626",
    # "https://telegram.me/lenskiy_vestni/29702",
]


async def main_vk():
    vk = GenerateScreenshotVk()
    for url in urls_vk:
        start_time = time.time()
        name = url.split("/")[-1]
        try:
            await vk.generate_screen_shot(url, f"screenshots/vk/{name}.png")
        except Exception as e:
            print(f"Screenshot {name} does not exists, {e}")
        end_time = time.time()
        print(f"Время выполнения: {end_time - start_time} секунд для {name}")


async def main_tg():
    tg = GenerateScreenshotTg()
    for url in urls_tg:
        start_time = time.time()
        name = url.split("/")[-1]
        try:
            await tg.generate_screen_shot(url, f"screenshots/tg/{name}.png")
        except Exception as e:
            print(f"Screenshot {name} does not exists, {e}")
        end_time = time.time()
        print(f"Время выполнения: {end_time - start_time} секунд для {name}")


async def main_ok():
    ok = GenerateScreenshotOk()
    for url in urls_ok:
        start_time = time.time()
        name = url.split("/")[-1]
        try:
            await ok.generate_screen_shot(url, f"screenshots/ok/{name}.png")
        except Exception as e:
            print(f"Screenshot {name} does not exists, {e}")
        end_time = time.time()
        print(f"Время выполнения: {end_time - start_time} секунд для {name}")


asyncio.run(main_ok())
