import re
import shutil
from urllib.parse import parse_qs, urlsplit

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
    if parsed.scheme != "https" or not is_hanatour(parsed.hostname) or parsed.username or parsed.password or parsed.port not in (None, 443):
        raise ValueError("하나투어 HTTPS 상품 주소를 확인해 주세요.")
    return parsed._replace(fragment="").geturl()


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
            page.set_default_timeout(30000)

            def allow_hanatour(route):
                request = route.request
                parsed = urlsplit(request.url)
                if parsed.scheme == "https" and is_hanatour(parsed.hostname) and request.resource_type not in {"image", "media", "font"}:
                    route.continue_()
                else:
                    route.abort()

            page.route("**/*", allow_hanatour)
            response = page.goto(address, wait_until="domcontentloaded", timeout=22000)
            if response and not response.ok:
                raise ValueError(f"하나투어 상품 페이지 응답 오류 (HTTP {response.status}).")
            return _read_schedule_tab(page, Error)
        except Error as error:
            raise ValueError("하나투어 상품 페이지 접속 또는 응답 대기 시간이 초과됐습니다.") from error
        finally:
            browser.close()


def _read_schedule_tab(page, browser_error):
    day_pattern = re.compile(r"^\s*\d{1,2}\s*일차\s+\d{1,2}[./-]\d{1,2}")
    try:
        page.get_by_text("여행일정", exact=True).filter(visible=True).first.click()
    except browser_error as error:
        raise ValueError("하나투어 상품 페이지에서 ‘여행일정’ 탭을 찾거나 열지 못했습니다.") from error
    day_links = page.get_by_role("link", name=day_pattern)
    try:
        day_links.nth(1).wait_for(state="visible")
    except browser_error as error:
        raise ValueError("하나투어 ‘여행일정’ 탭은 열렸지만 날짜별 일정이 표시되지 않았습니다.") from error
    summary = day_links.all_inner_texts()
    body = page.locator("body").inner_text()
    start = re.search(r"(?m)^\s*1\s*일차\s+\d{1,2}[./-]\d{1,2}", body)
    detail = body[start.start():] if start else ""
    return "\n\n".join(("\n".join(summary), detail[:26000])).strip()
