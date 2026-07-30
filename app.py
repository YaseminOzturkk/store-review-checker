from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from io import BytesIO
import re
import time
import xml.etree.ElementTree as ET
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

import pandas as pd
import requests
import streamlit as st
from google_play_scraper import app as google_play_app
from openpyxl.styles import Alignment, Font, PatternFill


REQUEST_TIMEOUT = 25
APPLE_REVIEW_MAX_PAGES = 10
APPLE_REVIEW_PAGE_SIZE = 50
APPLE_RETRY_ATTEMPTS = 4
APPLE_REQUEST_DELAY = 0.35
APP_STORE_REVIEW_LIMIT = 500
MAX_LINKS_PER_RUN = 200
MAX_WORKERS = 3
FALLBACK_COUNTRY = "us"
FALLBACK_LANGUAGE = "en"
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150 Safari/537.36"
)

OUTPUT_COLUMNS = [
    "Uygulama",
    "Platform",
    "Ülke",
    "Dil",
    "Yazılı Yorum Var mı?",
    "Yazılı Yorum Sayısı",
    "Sayı Kesin mi?",
    "Puanlama Sayısı",
    "Ortalama Puan",
    "Son Yorum Tarihi",
    "Uygulama Kimliği",
    "Durum",
    "Hata",
    "Girilen Link",
    "Mağaza Linki",
    "Kaynak",
    "Kontrol Zamanı (UTC)",
]


def clean_urls(raw: str) -> list[str]:
    """Her satırdan benzersiz bir mağaza bağlantısı çıkarır."""
    urls: list[str] = []
    seen: set[str] = set()

    for line in (raw or "").splitlines():
        value = line.strip().strip('"').strip("'")
        if not value or value in seen:
            continue
        seen.add(value)
        urls.append(value)

    return urls


def parse_store_url(url: str) -> dict[str, str | None]:
    """Google Play veya App Store bağlantısını ayrıştırır."""
    parsed = urlparse(url.strip())
    host = parsed.netloc.lower().split(":", 1)[0]
    query = parse_qs(parsed.query)

    if host in {"play.google.com", "www.play.google.com"}:
        package_id = (query.get("id") or [None])[0]
        if not package_id:
            raise ValueError("Google Play linkinde uygulama paket kimliği bulunamadı.")

        return {
            "platform": "Google Play",
            "identifier": package_id,
            "country": ((query.get("gl") or [None])[0] or "").lower() or None,
            "language": ((query.get("hl") or [None])[0] or "").lower() or None,
        }

    if host in {
        "apps.apple.com",
        "www.apps.apple.com",
        "itunes.apple.com",
        "www.itunes.apple.com",
    }:
        app_id_match = re.search(r"(?:^|/)id(\d+)(?:$|[/?])", parsed.path)
        if not app_id_match:
            app_id_match = re.search(r"[?&]id=(\d+)", url)
        if not app_id_match:
            raise ValueError("App Store linkinde sayısal uygulama kimliği bulunamadı.")

        parts = [unquote(part) for part in parsed.path.split("/") if part]
        country = (
            parts[0].lower()
            if parts and re.fullmatch(r"[A-Za-z]{2}", parts[0])
            else None
        )

        return {
            "platform": "App Store",
            "identifier": app_id_match.group(1),
            "country": country,
            "language": None,
        }

    raise ValueError("Yalnızca Google Play ve Apple App Store linkleri destekleniyor.")


def normalize_google_language(language: str) -> str:
    """Google Play hl değerini scraper'ın beklediği kısa dile dönüştürür."""
    return language.replace("_", "-").split("-", 1)[0].lower()


def check_google_play(url: str, parsed: dict[str, str | None]) -> dict[str, Any]:
    country = (parsed.get("country") or FALLBACK_COUNTRY).lower()
    language = normalize_google_language(parsed.get("language") or FALLBACK_LANGUAGE)
    app_id = str(parsed["identifier"])

    payload = google_play_app(app_id, lang=language, country=country)
    review_count = payload.get("reviews")
    rating_count = payload.get("ratings")
    score = payload.get("score")

    return {
        "Uygulama": payload.get("title"),
        "Platform": "Google Play",
        "Ülke": country.upper(),
        "Dil": language,
        "Yazılı Yorum Var mı?": (
            "Evet" if review_count is not None and int(review_count) > 0 else "Hayır"
        ),
        "Yazılı Yorum Sayısı": (
            int(review_count) if review_count is not None else None
        ),
        "Sayı Kesin mi?": "Evet" if review_count is not None else "Bilinmiyor",
        "Puanlama Sayısı": int(rating_count) if rating_count is not None else None,
        "Ortalama Puan": round(float(score), 2) if score is not None else None,
        "Son Yorum Tarihi": None,
        "Uygulama Kimliği": app_id,
        "Durum": "Başarılı",
        "Hata": None,
        "Girilen Link": url,
        "Mağaza Linki": payload.get("url") or url,
        "Kaynak": "google-play-scraper",
        "Kontrol Zamanı (UTC)": datetime.now(timezone.utc).isoformat(
            timespec="seconds"
        ),
    }


def apple_lookup(app_id: str, country: str) -> dict[str, Any]:
    response = requests.get(
        "https://itunes.apple.com/lookup",
        params={"id": app_id, "country": country},
        headers={"User-Agent": USER_AGENT},
        timeout=REQUEST_TIMEOUT,
    )
    response.raise_for_status()

    results = response.json().get("results", [])
    if not results:
        raise ValueError(f"Uygulama {country.upper()} App Store mağazasında bulunamadı.")

    return results[0]


def _apple_request_with_retry(url: str) -> requests.Response:
    """Apple yorum akışını hız sınırlarına karşı kontrollü biçimde çağırır."""
    last_error: Exception | None = None

    for attempt in range(APPLE_RETRY_ATTEMPTS):
        try:
            response = requests.get(
                url,
                headers={
                    "User-Agent": USER_AGENT,
                    "Accept": "application/atom+xml, application/xml, text/xml",
                },
                timeout=REQUEST_TIMEOUT,
            )
        except requests.RequestException as exc:
            last_error = exc
            if attempt == APPLE_RETRY_ATTEMPTS - 1:
                raise
            time.sleep(1.5 * (attempt + 1))
            continue

        if response.status_code == 200:
            return response

        if response.status_code in {403, 429, 500, 502, 503, 504}:
            retry_after = response.headers.get("Retry-After")
            try:
                wait_seconds = float(retry_after) if retry_after else 2 ** attempt
            except ValueError:
                wait_seconds = 2 ** attempt

            last_error = RuntimeError(
                f"Apple yorum akışı geçici olarak {response.status_code} döndürdü."
            )
            if attempt < APPLE_RETRY_ATTEMPTS - 1:
                time.sleep(max(wait_seconds, 1.0))
                continue

        response.raise_for_status()

    if last_error:
        raise last_error
    raise RuntimeError("Apple yorum akışına erişilemedi.")


def _parse_apple_review_entries(xml_content: bytes) -> list[ET.Element]:
    """Apple Atom XML akışındaki yazılı yorum kayıtlarını döndürür."""
    root = ET.fromstring(xml_content)
    atom_ns = "{http://www.w3.org/2005/Atom}"
    return list(root.findall(f"{atom_ns}entry"))


def _entry_text(entry: ET.Element, tag: str) -> str | None:
    atom_ns = "{http://www.w3.org/2005/Atom}"
    node = entry.find(f"{atom_ns}{tag}")
    if node is None or node.text is None:
        return None
    return node.text.strip() or None


def count_apple_reviews_xml(
    country: str,
    app_id: str,
    max_reviews: int = APP_STORE_REVIEW_LIMIT,
) -> tuple[int | None, bool | None, str | None, str | None]:
    """
    Apple müşteri yorumları Atom XML akışını sayfalar halinde tarar.

    Dönüş:
    - erişilebilen yazılı yorum sayısı (`None`: belirlenemedi)
    - sayının kesinliği (`None`: belirlenemedi)
    - son yorum tarihi
    - uyarı metni
    """
    seen: set[str] = set()
    last_date: str | None = None
    max_pages = min(
        APPLE_REVIEW_MAX_PAGES,
        max(1, (max_reviews + APPLE_REVIEW_PAGE_SIZE - 1) // APPLE_REVIEW_PAGE_SIZE),
    )

    for page in range(1, max_pages + 1):
        endpoint = (
            f"https://itunes.apple.com/{country}/rss/customerreviews/"
            f"page={page}/id={app_id}/sortby=mostrecent/xml"
        )

        try:
            response = _apple_request_with_retry(endpoint)
        except requests.HTTPError as exc:
            status = exc.response.status_code if exc.response is not None else None
            warning = (
                f"Apple yorum akışı HTTP {status} döndürdü; bu durum yorum olmadığı "
                "anlamına gelmez."
            )
            if seen:
                return len(seen), False, last_date, warning
            return None, None, None, warning
        except (requests.RequestException, ET.ParseError, RuntimeError) as exc:
            warning = f"Apple yorum akışı okunamadı: {exc}"
            if seen:
                return len(seen), False, last_date, warning
            return None, None, None, warning

        try:
            entries = _parse_apple_review_entries(response.content)
        except ET.ParseError as exc:
            warning = f"Apple yorum XML'i çözümlenemedi: {exc}"
            if seen:
                return len(seen), False, last_date, warning
            return None, None, None, warning

        if not entries:
            return len(seen), True, last_date, None

        new_items = 0
        for entry in entries:
            entry_id = _entry_text(entry, "id") or ET.tostring(
                entry, encoding="unicode"
            )
            if entry_id in seen:
                continue

            seen.add(entry_id)
            new_items += 1

            if last_date is None:
                last_date = _entry_text(entry, "updated")

            if len(seen) >= max_reviews:
                return len(seen), False, last_date, None

        if new_items == 0:
            return len(seen), True, last_date, None

        # Apple genellikle sayfa başına 50 yorum verir. Daha az kayıt son sayfadır.
        if len(entries) < APPLE_REVIEW_PAGE_SIZE:
            return len(seen), True, last_date, None

        time.sleep(APPLE_REQUEST_DELAY)

    return len(seen), False, last_date, None


def check_app_store(url: str, parsed: dict[str, str | None]) -> dict[str, Any]:
    country = (parsed.get("country") or FALLBACK_COUNTRY).lower()
    app_id = str(parsed["identifier"])
    metadata = apple_lookup(app_id, country)

    review_count, exact, last_date, review_warning = count_apple_reviews_xml(
        country=country,
        app_id=app_id,
    )

    rating_count = metadata.get("userRatingCount")
    score = metadata.get("averageUserRating")

    if review_count is None:
        has_written_reviews = "Bilinmiyor"
        review_count_display: int | str | None = None
        exact_display = "Bilinmiyor"
        status = "Kısmi"
    elif review_count > 0:
        has_written_reviews = "Evet"
        review_count_display = review_count if exact else f"{review_count}+"
        exact_display = "Evet" if exact else "Hayır"
        status = "Başarılı" if exact else "Kısmi"
    else:
        has_written_reviews = "Hayır" if exact else "Bilinmiyor"
        review_count_display = 0 if exact else None
        exact_display = "Evet" if exact else "Bilinmiyor"
        status = "Başarılı" if exact else "Kısmi"

    return {
        "Uygulama": metadata.get("trackName"),
        "Platform": "App Store",
        "Ülke": country.upper(),
        "Dil": None,
        "Yazılı Yorum Var mı?": has_written_reviews,
        "Yazılı Yorum Sayısı": review_count_display,
        "Sayı Kesin mi?": exact_display,
        "Puanlama Sayısı": int(rating_count) if rating_count is not None else None,
        "Ortalama Puan": round(float(score), 2) if score is not None else None,
        "Son Yorum Tarihi": last_date,
        "Uygulama Kimliği": app_id,
        "Durum": status,
        "Hata": review_warning,
        "Girilen Link": url,
        "Mağaza Linki": metadata.get("trackViewUrl") or url,
        "Kaynak": "Apple Lookup + Apple Reviews XML",
        "Kontrol Zamanı (UTC)": datetime.now(timezone.utc).isoformat(
            timespec="seconds"
        ),
    }


def error_result(url: str, error: Exception) -> dict[str, Any]:
    return {
        "Uygulama": None,
        "Platform": "Bilinmiyor",
        "Ülke": None,
        "Dil": None,
        "Yazılı Yorum Var mı?": "Bilinmiyor",
        "Yazılı Yorum Sayısı": None,
        "Sayı Kesin mi?": "Bilinmiyor",
        "Puanlama Sayısı": None,
        "Ortalama Puan": None,
        "Son Yorum Tarihi": None,
        "Uygulama Kimliği": None,
        "Durum": "Hata",
        "Hata": str(error),
        "Girilen Link": url,
        "Mağaza Linki": None,
        "Kaynak": None,
        "Kontrol Zamanı (UTC)": datetime.now(timezone.utc).isoformat(
            timespec="seconds"
        ),
    }


def check_one_link(url: str) -> dict[str, Any]:
    try:
        parsed = parse_store_url(url)
        if parsed["platform"] == "Google Play":
            return check_google_play(url, parsed)
        return check_app_store(url, parsed)
    except Exception as exc:  # Her linkin hatası kendi satırında gösterilir.
        return error_result(url, exc)


def check_links(urls: list[str], progress_bar: Any, status_box: Any) -> pd.DataFrame:
    results: list[dict[str, Any] | None] = [None] * len(urls)
    workers = min(MAX_WORKERS, len(urls))

    with ThreadPoolExecutor(max_workers=workers) as executor:
        future_map = {
            executor.submit(check_one_link, url): index
            for index, url in enumerate(urls)
        }

        completed = 0
        for future in as_completed(future_map):
            index = future_map[future]
            results[index] = future.result()
            completed += 1
            progress_bar.progress(completed / len(urls))
            status_box.caption(f"{completed}/{len(urls)} link kontrol edildi")

    return pd.DataFrame(
        [item for item in results if item is not None],
        columns=OUTPUT_COLUMNS,
    )


def dataframe_to_excel(frame: pd.DataFrame) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name="Store Kontrolü")
        worksheet = writer.sheets["Store Kontrolü"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        header_fill = PatternFill("solid", fgColor="E8EEF9")
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in column_cells)
            width = min(max(max_length + 2, 12), 55)
            worksheet.column_dimensions[column_cells[0].column_letter].width = width

    output.seek(0)
    return output.getvalue()


def render_results(frame: pd.DataFrame) -> None:
    successful = int((frame["Durum"] == "Başarılı").sum())
    partial = int((frame["Durum"] == "Kısmi").sum())
    with_reviews = int((frame["Yazılı Yorum Var mı?"] == "Evet").sum())
    errors = int((frame["Durum"] == "Hata").sum())

    metric_1, metric_2, metric_3, metric_4, metric_5 = st.columns(5)
    metric_1.metric("Toplam link", len(frame))
    metric_2.metric("Başarılı", successful)
    metric_3.metric("Kısmi", partial)
    metric_4.metric("Yorum bulunan", with_reviews)
    metric_5.metric("Hatalı", errors)

    st.dataframe(
        frame,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Girilen Link": st.column_config.LinkColumn("Girilen Link"),
            "Mağaza Linki": st.column_config.LinkColumn("Mağaza Linki"),
            "Ortalama Puan": st.column_config.NumberColumn(
                "Ortalama Puan", format="%.2f"
            ),
        },
    )

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    st.download_button(
        label="Excel Sonucunu İndir",
        data=dataframe_to_excel(frame),
        file_name=f"store-review-checker_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


st.set_page_config(
    page_title="Store Review Checker",
    page_icon="⭐",
    layout="wide",
)

st.title("⭐ Store Review Checker")
st.write(
    "Google Play ve App Store bağlantılarını satır satır yapıştırın. "
    "Uygulama ülke ve dil bilgilerini bağlantıdan otomatik okur."
)

raw_urls = st.text_area(
    "Google Play / App Store linkleri",
    height=320,
    placeholder=(
        "Her satıra bir bağlantı yapıştırın.\n\n"
        "https://play.google.com/store/apps/details?id=com.example.app&hl=en&gl=US\n"
        "https://apps.apple.com/us/app/example/id123456789"
    ),
)

check_button = st.button(
    "Linkleri Kontrol Et",
    type="primary",
    use_container_width=True,
)

if check_button:
    urls = clean_urls(raw_urls)

    if not urls:
        st.error("En az bir Google Play veya App Store linki girin.")
    elif len(urls) > MAX_LINKS_PER_RUN:
        st.error(f"Tek seferde en fazla {MAX_LINKS_PER_RUN} link kontrol edilebilir.")
    else:
        progress_bar = st.progress(0)
        status_box = st.empty()

        with st.spinner("Mağaza bilgileri kontrol ediliyor..."):
            result_frame = check_links(urls, progress_bar, status_box)

        progress_bar.empty()
        status_box.empty()
        st.session_state["store_results"] = result_frame

if "store_results" in st.session_state:
    st.divider()
    render_results(st.session_state["store_results"])

with st.expander("Nasıl çalışır?"):
    st.markdown(
        """
- Google Play bağlantılarında `gl` ülke, `hl` dil kodu olarak kullanılır.
- App Store bağlantılarında `/tr/`, `/us/`, `/gb/` gibi bölüm ülke mağazasını belirler.
- Linkte ülke veya dil bulunmazsa teknik yedek olarak `US / en` kullanılır.
- App Store yazılı yorumları Apple müşteri yorumları XML akışından en fazla 500 kayda kadar taranır.
- Sınırın aşılması halinde yorum sayısı `500+` biçiminde gösterilir.
- Apple akışı geçici olarak 403/429/404 döndürürse sonuç yanlış biçimde “Hayır” olmaz; “Bilinmiyor” veya “Kısmi” gösterilir.
- Puanlama sayısı ile yazılı yorum sayısı ayrı sütunlardır.
        """
    )
