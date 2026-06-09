import os
import sys
import re
from datetime import datetime, timedelta
import urllib.parse
import pandas as pd
import openpyxl
from playwright.sync_api import sync_playwright

# --- Configuration & Rules ---
KEYWORDS = ["devops", "sre"]
START_URL = "https://europa.eu/eures/portal/jv-se/search?page=1&resultsPerPage=50&orderBy=BEST_MATCH&locationCodes=de&keywordsEverywhere=devops%2Bsre&experience=NS&positionScheduleCodes=fulltime&publicationPeriod=LAST_THREE_DAYS&requiredLanguages=en(C2),de(C2)&previousPageType=findJob&lang=en"
PAGE_GROUP = int(os.getenv("PAGE_GROUP", "1"))
PAGES_PER_RUN = 10
START_PAGE = (PAGE_GROUP - 1) * PAGES_PER_RUN + 1
END_PAGE = START_PAGE + PAGES_PER_RUN - 1
OUTPUT_FILENAME = f"EURES_DevOps_SRE_Jobs_Germany_pages_{START_PAGE}-{END_PAGE}.xlsx"
OUTPUT_HTML_FILENAME = f"EURES_DevOps_SRE_Jobs_Germany_pages_{START_PAGE}-{END_PAGE}.html"
APPLIED_JOBS_FILE = os.getenv("APPLIED_JOBS_FILE", "applied_jobs.txt")
SAVE_APPLIED = os.getenv("SAVE_APPLIED", "0") == "1"


def parse_eures_date(date_str):
    """Parses typical EURES date formats into a standard datetime object."""
    if not date_str:
        return None
    normalized = date_str.strip().lower()
    now = datetime.now()

    if "today" in normalized:
        return now
    if "yesterday" in normalized:
        return now - timedelta(days=1)

    days_ago = re.search(r"(\d+)\s+days?\s+ago", normalized)
    if days_ago:
        return now - timedelta(days=int(days_ago.group(1)))

    hours_ago = re.search(r"(\d+)\s+hours?\s+ago", normalized)
    if hours_ago:
        return now - timedelta(hours=int(hours_ago.group(1)))

    try:
        # Standard YYYY-MM-DD or DD/MM/YYYY
        if "-" in normalized:
            return datetime.strptime(normalized[:10], "%Y-%m-%d")
        elif "/" in normalized:
            return datetime.strptime(normalized[:10], "%d/%m/%Y")
    except Exception:
        pass
    return None

def extract_apply_link(page):
    """
    Locates the 'How to apply' section in the job detail page and returns the first valid
    URL or email found in that section.
    """
    try:
        page.wait_for_selector('dt:has-text("How to apply"), h2:has-text("How to apply")', timeout=10000)

        apply_dd = None
        if page.locator('dt:has-text("How to apply")').count() > 0:
            apply_dd = page.locator('xpath=//dt[contains(normalize-space(.),"How to apply")]/following-sibling::dd[1]').first
        elif page.locator('h2:has-text("How to apply")').count() > 0:
            apply_dd = page.locator('xpath=//h2[contains(normalize-space(.),"How to apply")]/following::dd[1]').first

        if apply_dd and apply_dd.count() > 0:
            link_elems = apply_dd.locator('a[href]').all()
            hrefs = [l.get_attribute('href') for l in link_elems if l.get_attribute('href')]
            valid_hrefs = [h for h in hrefs if h and h.startswith(('http://', 'https://'))]
            if valid_hrefs:
                return valid_hrefs[0]

            email_links = apply_dd.locator('a[href^="mailto:"]').all()
            if email_links:
                email_href = email_links[0].get_attribute('href')
                return email_href.replace('mailto:', '').split('?')[0]

            text_content = apply_dd.text_content() or ''
            url_match = re.search(r'https?://\S+', text_content)
            if url_match:
                return url_match.group(0).rstrip('.,;')

        return 'Apply via EURES'
    except Exception:
        return 'Apply via EURES'

def build_target_url(page_num):
    parsed_url = urllib.parse.urlparse(START_URL)
    query_params = urllib.parse.parse_qs(parsed_url.query)
    query_params['resultsPerPage'] = ['50']
    query_params['page'] = [str(page_num)]
    new_query = urllib.parse.urlencode(query_params, doseq=True)
    return urllib.parse.urlunparse(parsed_url._replace(query=new_query))


def load_applied_jobs():
    applied = set()
    if not os.path.exists(APPLIED_JOBS_FILE):
        return applied

    try:
        with open(APPLIED_JOBS_FILE, 'r', encoding='utf-8') as source:
            for raw_line in source:
                line = raw_line.strip()
                if not line or line.startswith('#'):
                    continue
                applied.add(line)
                if '/jv-details/' in line:
                    matched = re.search(r'/jv-details/([^?/#]+)', line)
                    if matched:
                        applied.add(matched.group(1))
                elif line.startswith('http') and 'jv-details' not in line:
                    matched = re.search(r'/jv-details/([^?/#]+)', line)
                    if matched:
                        applied.add(matched.group(1))
    except Exception:
        print(f"[!] Warning: could not read applied jobs file '{APPLIED_JOBS_FILE}'")
    return applied


def save_applied_jobs(job_urls):
    if not SAVE_APPLIED:
        return
    try:
        with open(APPLIED_JOBS_FILE, 'a', encoding='utf-8') as sink:
            for job_url in job_urls:
                sink.write(f"{job_url}\n")
    except Exception:
        print(f"[!] Warning: could not append to applied jobs file '{APPLIED_JOBS_FILE}'")


def scrape_eures():
    all_jobs = []
    extraction_errors = []
    unopened_jobs = []
    skipped_applied = 0
    applied_jobs = load_applied_jobs()

    print(f"[!] Initiating scraping engine for pages {START_PAGE} through {END_PAGE} (PAGE_GROUP={PAGE_GROUP})")
    print(f"[!] Applied jobs loaded: {len(applied_jobs)} from {APPLIED_JOBS_FILE}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        page = context.new_page()
        detail_page = context.new_page()

        for current_page_num in range(START_PAGE, END_PAGE + 1):
            target_url = build_target_url(current_page_num)
            print(f"[*] Processing Search Result Page #{current_page_num}: {target_url}")
            page.goto(target_url, wait_until="domcontentloaded", timeout=60000)

            try:
                page.wait_for_selector("[role='article'][aria-label='Job vacancy result']", timeout=30000)
            except Exception:
                print(f"[-] Page {current_page_num} had no job listings or timed out.")
                break

            job_cards = page.locator("[role='article'][aria-label='Job vacancy result']")
            job_count = job_cards.count()
            if job_count == 0:
                print(f"[-] Page {current_page_num} contained zero jobs.")
                break

            print(f"[*] Page {current_page_num} contains {job_count} listings")
            for index in range(job_count):
                card = job_cards.nth(index)
                try:
                    title_el = card.locator("a.ecl-link--standalone").first
                    title = title_el.text_content().strip() if title_el.count() > 0 else ""
                    href = title_el.get_attribute("href") if title_el.count() > 0 else ""
                    job_url = urllib.parse.urljoin(page.url, href) if href else ""

                    company_el = card.locator("li[id^='jv-employer-name']").first
                    company = company_el.text_content().strip() if company_el.count() > 0 else "Unknown Employer"

                    location_el = card.locator("span[id^='location']").first
                    location_text = location_el.text_content().strip() if location_el.count() > 0 else "Germany"
                    location_text = re.sub(r"\s+:\s+", ", ", location_text)
                    city = location_text.replace("Germany", "").strip().strip(':').strip()
                    if not city:
                        city = "Germany"

                    date_el = card.locator("li[id^='date']").first
                    date_text = date_el.text_content().strip() if date_el.count() > 0 else ""
                    match = re.search(r"(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d+\s+days?\s+ago|\d+\s+hours?\s+ago|today|yesterday)", date_text, re.IGNORECASE)
                    if match:
                        date_text = match.group(1)

                    schedule_el = card.locator("span[id^='position-schedule']").first
                    schedule = schedule_el.text_content().strip() if schedule_el.count() > 0 else ""

                    description_el = card.locator(".ecl-content-block__description").first
                    summary = description_el.text_content().strip() if description_el.count() > 0 else ""

                    job_id = None
                    if job_url:
                        matched_job_id = re.search(r'/jv-details/([^?/#]+)', job_url)
                        job_id = matched_job_id.group(1) if matched_job_id else None

                    print(f"    [job {index + 1}/{job_count}] title='{title}' url={job_url}")

                    if job_url and (job_url in applied_jobs or (job_id and job_id in applied_jobs)):
                        skipped_applied += 1
                        print(f"      -> skipped already applied")
                        continue

                    apply_link = "Apply via EURES"
                    if job_url:
                        try:
                            detail_page.goto(job_url, wait_until="domcontentloaded", timeout=30000)
                            apply_link = extract_apply_link(detail_page)
                        except Exception as e:
                            extraction_errors.append(f"Error loading detail page for {job_url}: {str(e)}")

                    if title:
                        all_jobs.append({
                            "Job Title": title,
                            "Company": company,
                            "Country": "Germany",
                            "Location": city,
                            "Schedule": schedule,
                            "Date Posted": date_text,
                            "Summary": summary,
                            "Job URL": job_url,
                            "Apply Link": apply_link
                        })
                    else:
                        unopened_jobs.append(job_url)
                except Exception as e:
                    extraction_errors.append(f"Error extracting listing card #{index + 1} on page {current_page_num}: {str(e)}")

        browser.close()

    print(f"[!] Skipped already applied jobs: {skipped_applied}")
    return all_jobs, extraction_errors, unopened_jobs

def build_excel_report(jobs, extraction_errors, unopened_jobs):
    now = datetime.now()
    
    # Structural Data Clean and Sort Pre-processing
    df_raw = pd.DataFrame(jobs)
    if df_raw.empty:
        print("[!] Execution stopped: Extraction returned 0 records.")
        return
        
    df_raw['Parsed Date'] = df_raw['Date Posted'].apply(parse_eures_date)
    # Strip records missing vital schema metrics
    df_raw = df_raw.dropna(subset=['Job Title', 'Parsed Date', 'Job URL'])
    df_raw = df_raw.sort_values(by='Parsed Date', ascending=False)
    
    # Apply Time-Horizon logic masks
    mask_1_day = df_raw['Parsed Date'] >= (now - timedelta(days=1))
    mask_3_days = df_raw['Parsed Date'] >= (now - timedelta(days=3))
    mask_7_days = df_raw['Parsed Date'] >= (now - timedelta(days=7))
    
    # Process distinct segments with clean primary de-duplication rules
    df_all = df_raw.drop_duplicates(subset=['Job URL']).drop(columns=['Parsed Date'])
    df_1 = df_raw[mask_1_day].drop_duplicates(subset=['Job URL'], keep='first').drop(columns=['Parsed Date'])
    df_3 = df_raw[mask_3_days].drop_duplicates(subset=['Job URL'], keep='first').drop(columns=['Parsed Date'])
    df_7 = df_raw[mask_7_days].drop_duplicates(subset=['Job URL'], keep='first').drop(columns=['Parsed Date'])
    
    # Re-verify baseline dataset lengths
    total_unique_jobs = len(df_all)
    missing_direct_apply = len(df_raw[df_raw['Apply Link'] == "Apply via EURES"].drop_duplicates(subset=['Job URL']))
    
    # Summary Table Construction
    summary_metrics = {
        "Metric": [
            "Total Jobs Found", 
            "Jobs Last 1 Day", 
            "Jobs Last 3 Days", 
            "Jobs Last 7 Days", 
            "Jobs Missing Direct Apply Link"
        ],
        "Value": [
            total_unique_jobs, 
            len(df_1), 
            len(df_3), 
            len(df_7), 
            missing_direct_apply
        ]
    }
    df_summary = pd.DataFrame(summary_metrics)
    
    # List of jobs defaulting back to internal system routing
    df_missing_links_list = df_raw[df_raw['Apply Link'] == "Apply via EURES"].drop_duplicates(subset=['Job URL']).drop(columns=['Parsed Date'])

    # Write out structural contents to Excel target
    with pd.ExcelWriter(OUTPUT_FILENAME, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Summary', index=False)
        df_missing_links_list.to_excel(writer, sheet_name='Summary', startrow=8, index=False)
        
        df_all.to_excel(writer, sheet_name='All_Jobs', index=False)
        df_1.to_excel(writer, sheet_name='Last_1_Day', index=False)
        df_3.to_excel(writer, sheet_name='Last_3_Days', index=False)
        df_7.to_excel(writer, sheet_name='Last_7_Days', index=False)
        
        # Style Sheets to comply with Workbook Visual Design requirements
        workbook = writer.book
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            
            # Freeze First Row
            worksheet.freeze_panes = "A2"
            
            # Format Headers and Convert Links to Active Hyperlinks
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row), start=1):
                for col_idx, cell in enumerate(row, start=1):
                    # Style Header Element explicitly
                    if row_idx == 1:
                        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
                        cell.fill = openpyxl.styles.PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                    else:
                        # Auto-wrap string elements checking for Apply Link or Job URL configurations
                        header_val = worksheet.cell(row=1, column=col_idx).value
                        if header_val in ["Apply Link", "Job URL"] and cell.value and str(cell.value).startswith("http"):
                            cell.hyperlink = cell.value
                            cell.font = openpyxl.styles.Font(color="0000FF", underline="single")
            
            # Auto-size all columns based on content length bounding
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Output Console Diagnostic metrics report
    print("\n" + "="*50)
    print("       EURES SCRAPER EXECUTION SUMMARY")
    print("="*50)
    print(f"Total Jobs Found:                {total_unique_jobs}")
    print(f"Number of jobs in Last_1_Day:   {len(df_1)}")
    print(f"Number of jobs in Last_3_Days:  {len(df_3)}")
    print(f"Number of jobs in Last_7_Days:  {len(df_7)}")
    print(f"Excel file path:                {os.path.abspath(OUTPUT_FILENAME)}")
    print(f"HTML file path:                 {os.path.abspath(OUTPUT_HTML_FILENAME)}")
    print(f"Any extraction errors:          {len(extraction_errors)}")
    print(f"Any jobs that couldn't open:    {len(unopened_jobs)}")
    print("="*50 + "\n")
    
    if extraction_errors:
        print("[!] Logging Encountered Extraction Errors:")
        for err in extraction_errors[:5]: # Cap output display logs
            print(f" -> {err}")

    # Write HTML fallback export with clickable links
    df_html = df_all.copy()
    for col in ["Job URL", "Apply Link"]:
        if col in df_html.columns:
            df_html[col] = df_html[col].apply(lambda x: f'<a href="{x}" target="_blank">{x}</a>' if pd.notna(x) and str(x).startswith("http") else x)
    df_html.to_html(OUTPUT_HTML_FILENAME, index=False, escape=False)

if __name__ == "__main__":
    jobs, errors, unopened = scrape_eures()
    build_excel_report(jobs, errors, unopened)