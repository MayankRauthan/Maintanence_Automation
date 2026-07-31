from playwright.sync_api import Playwright, sync_playwright



def get_data(context, link: str):
        page = context.new_page()

        # open base page (important if popup is triggered from there)
        page.goto(link, wait_until="domcontentloaded")

        # 👇 capture popup
        with page.expect_popup() as popup_info:
            page.goto(link)   # or click action if popup comes from click

        popup = popup_info.value

        popup.wait_for_load_state("domcontentloaded")
        popup.wait_for_timeout(5000)

        # Now work on popup, NOT page
        workflow = popup.locator(".sticky-footer .workflow")
        items = workflow.locator("> *")

       
        count =items.count()
        for i in range(count-1,-1,-1):
            item = items.nth(i).get_attribute("id")
            if item:
                page.close()
                popup.close()
                print("Status",item)
                return item
        page.close()

        
        return "ERROR"
# print(
#     get_data(
#         "http://machx.equant.com/machx/login.aspx?ApplicationExt=MAIL&MitroRO=true&CaseRefExt=XSF202603-00015"
#     )
# )
