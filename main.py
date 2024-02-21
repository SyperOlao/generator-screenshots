import asyncio
from time import sleep
from pyppeteer import launch

url = "https://vk.com/styd.pozor?w=wall-71729358_13128057"
url = "https://vk.com/instasamka?w=wall-182863061_923576"
# url = "https://vk.com/instasamka?w=wall-182863061_925645"
# url = "https://vk.com/al_feed.php?w=wall-90802042_128826"
#url = "https://vk.com/public179350392?w=wall-179350392_9"


async def delete_element_by_id(id: str, page):
    await page.evaluate(f'document.getElementById("{id}")?.remove();')

async def delete_element_by_class_name(class_name: str, page):
    await page.evaluate(f"""() => {{
        var elements = document.getElementsByClassName("{class_name}");
        if(!elements) return;
        while(elements?.length >  0){{
            elements[0]?.parentNode?.removeChild(elements[0]);
        }}
    }}""")

# wl_replies_block_wrap
async def main():
    browser = await launch()
    page = await browser.newPage()
    
    # await page.goto('https://vk.com')

    # # Fill in the username and password
    # await page.type('#index_email', '77473913755')
    # await page.keyboard.press('Enter')
    
    # element = await page.querySelector('[name="password"]')
    # await element.type('RG65hfhsllf65jfsh326')
    # await page.keyboard.press('Enter')


    # Wait for navigation to the desired page
    # await page.waitForNavigation()
    # await page.setViewport({'width':  1280, 'height':  800})

    await page.goto(url, {
    'waitUntil': 'load',
    'timeout':  300000  # Remove the timeout
})

    
    await page.waitForSelector('#wl_post')
      
    await delete_element_by_class_name("wl_replies_block_wrap", page)
    await delete_element_by_class_name("wl_replies_wrap", page)
    await delete_element_by_class_name("wl_replies", page)

    await delete_element_by_id("wl_replies_wrap", page)
    await delete_element_by_id("page_bottom_banners_root", page)

    
        
    element = await page.querySelector('#wk_content')
    
    
   
    if element:
        bounding_box = await element.boundingBox()
        print(bounding_box)
        height =  bounding_box['height']
    

    clip = {
        'width': int(bounding_box['width']),
        'height': int(bounding_box['height'] + 24)
    }
    await page.setViewport(clip)
   
            # 'x': int(bounding_box['x']),
            # 'y': int(bounding_box['y']),
    
    clip_1 = {
            'x': 0,
            'y': 24,
            'width': int(bounding_box['width']),
            'height': int(bounding_box['height'])
        }
    # 'clip': clip
    sleep(1)
    await page.screenshot({'path': 'test.png', 'clip': clip_1})
    await browser.close()


asyncio.run(main())


# wl_replies_wrap
