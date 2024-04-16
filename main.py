import asyncio

from pyppeteer.errors import NetworkError

from generate_screenshot.ok.generate_ok import GenerateScreenshotOk
from generate_screenshot.telegram.generate_tg import GenerateScreenshotTg
from generate_screenshot.vk.generate_vk import GenerateScreenshotVk
from config.config import config_env, logger

urls_vk_2 = [
    "http://vk.com/wall-57288440_4675150",
    "http://vk.com/wall-57288440_4676035",
    "http://vk.com/wall-61067996_1799536",
    "http://vk.com/wall-151168449_426502",
    "http://vk.com/wall-71138052_14043",
    "http://vk.com/wall-144865761_78152",
    "http://vk.com/wall-144865761_77700",
    "http://vk.com/wall-158810690_36241",
    "http://vk.com/wall-178517743_19218",
    "http://vk.com/wall-2264907_966",
    "http://vk.com/wall-178517743_19307",
    "http://vk.com/wall-158810690_35849",
    "http://vk.com/wall-185808385_1114025",
    "http://vk.com/wall-170916587_160232",
    "http://vk.com/wall-100098188_415599",
    "http://vk.com/wall678932830_60695",
    "http://vk.com/wall678932830_60796",
    "http://vk.com/wall625204288_5398",
    "http://vk.com/wall-211635463_959",
    "http://vk.com/wall-2264907_1001",
    "http://vk.com/wall-186255553_6774",
    "http://vk.com/wall-6685947_2826",
    "http://vk.com/wall-154371296_7879",
    "http://vk.com/wall-4734673_4767",
    "http://vk.com/wall-193815220_831",
    "http://vk.com/wall-208865227_2160",
    "http://vk.com/wall-161599180_1488",
]

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
        await vk.login_vk(config_env['VK_LOGIN'], config_env['VK_PASSWORD'])
        await vk.generate_screen_shots(urls_vk_2, "screenshots/vk")
    except NetworkError as e:
        logger.error(f"NetworkError: {e}")
    # await vk.browser_close()


async def main_tg():
    tg = GenerateScreenshotTg()
    await tg.generate_screen_shots(urls_tg, f"screenshots/tg")


async def main_ok():
    ok = GenerateScreenshotOk()
    await ok.generate_screen_shots(urls_tg, f"screenshots/ok")


asyncio.run(main_vk())
