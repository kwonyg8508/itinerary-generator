import re
import shutil
from urllib.parse import parse_qs, urlencode, urlsplit

PRODUCT_PATH = "/trp/pkg/CHPC0PKG0200M200"


def is_hanatour(host):
    return bool(host and (host == "hanatour.com" or host.endswith(".hanatour.com")))


def product_url(url):
    parsed = urlsplit(url)
    query = parse_qs(parsed.query)
    product = query.get("pkgCd", [""])[0]
    departure = query.get("depDay", [""])[0]
    if (
        parsed.path != PRODUCT_PATH
        or not re.fullmatch(r"[A-Z0-9]{8,24}", product)
        or (departure and not re.fullmatch(r"\d{8}", departure))
    ):
        raise ValueError("하나투어 상품 상세 주소와 상품코드(pkgCd)를 확인해 주세요.")
    params = {"pkgCd": product}
    if departure:
        params["depDay"] = departure
    params["newWin"] = "true"
    return f"https://www.hanatour.com{PRODUCT_PATH}?{urlencode(params)}"


def read_hanatour_itinerary(url):
    from playwright.sync_api import Error, sync_playwright

    address = product_url(url)
    chromium = shutil.which("chromium") or shutil.which("chromium-browser")
    if not chromium:
        raise ValueError("하나투어 일정 읽기에 필요한 Chromium이 설치되지 않았습니다. packages.txt를 확인해 주세요.")

    with sync_playwright() as playwright:
        try:
            browser = playwright.chromium.launch(executable_path=chromium, headless=True)
        except Error as error:
            raise ValueError("일정 읽기용 Chromium을 시작하지 못했습니다. Streamlit 배포 로그를 확인해 주세요.") from error
        try:
            page = browser.new_page(locale="ko-KR", service_workers="block")
            page.set_default_timeout(18000)

            def allow_hanatour(route):
                request = route.request
                parsed = urlsplit(request.url)
                if parsed.scheme == "https" and is_hanatour(parsed.hostname) and request.resource_type not in {"image", "media", "font"}:
                    route.continue_()
                else:
                    route.abort()

            page.route("**/*", allow_hanatour)
            page.goto(address, wait_until="domcontentloaded", timeout=22000)
            return _read_schedule_tab(page)
        except Error as error:
            raise ValueError("하나투어 상품 화면에서 여행일정 탭의 일자별 내용을 불러오지 못했습니다.") from error
        finally:
            browser.close()


def _read_schedule_tab(page):
    page.get_by_role("link", name="여행일정", exact=True).click()
    second_day = page.get_by_role("link", name=re.compile(r"^2일차\s+\d{2}/\d{2}")).first
    second_day.wait_for(state="visible")
    day_links = page.get_by_role("link", name=re.compile(r"^\d{1,2}일차\s+\d{2}/\d{2}"))
    summary = day_links.all_inner_texts()
    if len(summary) < 2:
        raise ValueError("하나투어에서 일자별 일정이 충분히 표시되지 않았습니다.")
    body = page.locator("body").inner_text()
    start = re.search(r"(?m)^\s*1일차\s+\d{2}/\d{2}", body)
    detail = body[start.start():] if start else ""
    return "\n\n".join(("\n".join(summary), detail[:26000])).strip()
