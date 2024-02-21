import asyncio
from generate_screenshot.vk.generate_vk import GenerateScreenshotVk

url = "https://vk.com/styd.pozor?w=wall-71729358_13128057"
# url = "https://vk.com/instasamka?w=wall-182863061_923576"
# url = "https://vk.com/instasamka?w=wall-182863061_925645"
# url = "https://vk.com/al_feed.php?w=wall-90802042_128826"
# url = "https://vk.com/istoriyadonbassa?w=wall-189776399_233201 "
# url = "https://vk.com/public179350392?w=wall-179350392_9"
# url = "https://vk.com/istoriyadonbassa?w=wall-189776399_233200"
# url = "https://vk.com/delphine_chik?w=wall515103149_134"
url = "https://vk.com/@-194707757-rss-1610825261-1959186610#"
#url = "https://vk.com/wall-141359201_1489951"
# url = "https://vk.com/wall-93250065_1198328"




# wl_replies_block_wrap
async def main():
   vk = GenerateScreenshotVk()
   await vk.generate_screen_shot(url, "test2.png")


asyncio.run(main())


# wl_replies_wrap
