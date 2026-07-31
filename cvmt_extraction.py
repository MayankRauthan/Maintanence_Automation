from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import re
 
 
def safe_inner_text(locator, timeout=2000):
    """Return inner_text safely without crashing on timeout."""
    try:
        return locator.inner_text(timeout=timeout)
    except PlaywrightTimeoutError:
        return ""
    except Exception:
        return ""
 
 
def extract_status_from_container(container):
    """
    container can be Page or Frame
    Returns status string or None
    """
 
    # 1) Direct status labels
    try:
        labels = container.locator("label[for='main_status']")
        count = labels.count()
        if count > 0:
            values = []
            for i in range(count):
                txt = safe_inner_text(labels.nth(i), timeout=1500).strip()
                if txt:
                    values.append(txt)
 
            cleaned = [
                v for v in values
                if v.lower().replace(" ", "") not in ("status:", "status")
            ]
            if cleaned:
                return cleaned[0]
            if values:
                return values[-1]
    except Exception:
        pass
 
    # 2) Table cell containing Status:
    try:
        cell = container.locator("td:has-text('Status:')").first
        if cell.count() > 0:
            text = safe_inner_text(cell, timeout=1500).replace("\xa0", " ").strip()
            m = re.search(r"Status:\s*(.+)", text, re.IGNORECASE)
            if m and m.group(1).strip():
                return m.group(1).strip()
    except Exception:
        pass
 
    # 3) Very safe page content fallback (faster than waiting on body locator)
    try:
        html = container.content()
        text = re.sub(r"<[^>]+>", " ", html)  # strip tags roughly
        text = re.sub(r"\s+", " ", text).replace("\xa0", " ")
        m = re.search(r"Status:\s*([A-Za-z][A-Za-z0-9 _-]+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    except Exception:
        pass
 
    return None
 
 
def get_data(context, link: str):
        page = context.new_page()
 
        page.goto(link, wait_until="domcontentloaded", timeout=80000)
 
        # If login/manual SSO is needed, give some time
        page.wait_for_timeout(6000)
 
        status_value = None
 
        # 1) Try mainFrame first (based on your inspect)
        main_frame = page.frame(name="mainFrame")
        if main_frame:
            try:
                main_frame.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
            status_value = extract_status_from_container(main_frame)
 
        # 2) Try page itself
        if not status_value:
            status_value = extract_status_from_container(page)
 
        # 3) Try all frames
        if not status_value:
            for fr in page.frames:
                try:
                    s = extract_status_from_container(fr)
                    if s:
                        status_value = s
                        break
                except Exception:
                    continue
        page.close()
        return status_value
 
 
# if __name__ == "__main__":
#     test_link = "https://cmt.sso.infra.ftgroup/login?maintRef=REF/2026.13779"
#     result = get_data(test_link)
#     print(result)
 