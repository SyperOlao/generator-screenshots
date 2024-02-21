from pyppeteer import launch


class GenerateScreenshot:
    def __init__(self):
        self.browser = None
        self.page = None
        

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


    async def _delete_by_class_names(self, class_names):
        for class_name in class_names:
            await self._delete_element_by_class_name(class_name)


    async def _delete_by_ids(self, ids):
        for id in ids:
            await self._delete_element_by_id(id)


    async def _check_existing_element_by_id(self, id: str):
        element = await self.page.querySelector(
            f"{id}",
        )
        return element is not None


    async def _get_existing_element(self, ids):
        for id in ids:
            if await self._check_existing_element_by_id(id):
                return id
        return None
