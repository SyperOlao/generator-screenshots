import asyncio

from pyppeteer.errors import NetworkError

from generate_screenshot.ok.generate_ok import GenerateScreenshotOk
from generate_screenshot.telegram.generate_tg import GenerateScreenshotTg
from generate_screenshot.vk.generate_vk import GenerateScreenshotVk
from config.config import config_env, logger

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
    "https://ok.ru/group/62848965345443/topic/155617158681507",
    "https://ok.ru/group/70000001070863/topic/156976990130191",
    "https://ok.ru/group/70000002276393/topic/157180683767081",
    "https://ok.ru/group/70000000716428/topic/157205056791948",
    "https://ok.ru/group/70000001209590/topic/157095483144438",
    "https://ok.ru/group/70000001125216/topic/155673607039072",
    "https://ok.ru/group/70000001155735/topic/156193421513623",
    "https://ok.ru/group/70000004708458/topic/156795335883626",
    "https://telegram.me/lenskiy_vestni/29702",
]


async def main_vk():
    vk = GenerateScreenshotVk()
    await vk.browser_open()
    try:
      #  await vk.login_vk(config_env['VK_LOGIN'], config_env['VK_PASSWORD'])
        await vk.generate_screen_shots(urls_vk, "screenshots/vk")
    except NetworkError as e:
        logger.error(f"NetworkError: {e}")
    await vk.browser_close()


async def main_tg():
    tg = GenerateScreenshotTg()
    await tg.generate_screen_shots(urls_tg, f"screenshots/tg")


async def main_ok():
    ok = GenerateScreenshotOk()
    await ok.generate_screen_shots(urls_tg, f"screenshots/ok")


asyncio.run(main_vk())
