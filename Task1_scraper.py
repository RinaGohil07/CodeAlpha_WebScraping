# ==============================================================
#   CodeAlpha Internship | Task 1: Web Scraping
#   Project  : Quotes Data Scraper & Analysis
#   Website  : quotes.toscrape.com (beginner-friendly practice site)
#   Author   : (Your Name)
#   Libraries: requests, BeautifulSoup, pandas, matplotlib
# ==============================================================
#
#  WHAT THIS SCRIPT DOES:
#    1. Scrapes all quotes, authors, and tags from the website
#    2. Cleans and organizes the data using pandas
#    3. Saves everything to CSV and Excel files
#    4. Creates 6 professional charts for analysis
#
#  HOW TO RUN:
#    pip install requests beautifulsoup4 pandas matplotlib seaborn openpyxl
#    python scraper.py
# ==============================================================

import requests
from bs4 import BeautifulSoup
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import seaborn as sns
from collections import Counter
import time

# ── SETTINGS ─────────────────────────────────────────────────
BASE_URL     = "https://quotes.toscrape.com"
OUTPUT_CSV   = "quotes_data.csv"
OUTPUT_EXCEL = "quotes_data.xlsx"
OUTPUT_CHART = "quotes_analysis.png"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}
# ─────────────────────────────────────────────────────────────


# =============================================================
#  STEP 1 — SCRAPE DATA
# =============================================================

def scrape_all_quotes():
    """
    Scrapes all quotes from quotes.toscrape.com.
    Navigates through all pages automatically using the 'Next' button.
    Returns a list of dictionaries with quote details.
    """
    all_quotes = []
    url = BASE_URL
    page_num = 1

    print("\n" + "="*55)
    print("  Starting Web Scraper...")
    print("="*55)

    while url:
        print(f"\n  Scraping Page {page_num}: {url}")

        try:
            response = requests.get(url, headers=HEADERS, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            # ── Find all quote blocks on this page ──────────
            quote_divs = soup.find_all("div", class_="quote")

            for div in quote_divs:
                # Quote Text (remove surrounding quotes)
                text_tag = div.find("span", class_="text")
                text = text_tag.get_text(strip=True) if text_tag else ""
                text = text.strip('\u201c\u201d"')          # remove curly quotes

                # Author Name
                author_tag = div.find("small", class_="author")
                author = author_tag.get_text(strip=True) if author_tag else "Unknown"

                # Author Bio Link
                author_link_tag = div.find("a", href=True)
                author_link = BASE_URL + author_link_tag["href"] if author_link_tag else ""

                # Tags (can be multiple per quote)
                tag_elements = div.find_all("a", class_="tag")
                tags = [t.get_text(strip=True) for t in tag_elements]

                # Extra metrics we calculate ourselves
                word_count  = len(text.split())
                char_count  = len(text)
                tag_count   = len(tags)

                all_quotes.append({
                    "Quote"       : text,
                    "Author"      : author,
                    "Tags"        : ", ".join(tags),
                    "Tag Count"   : tag_count,
                    "Word Count"  : word_count,
                    "Char Count"  : char_count,
                    "Page Number" : page_num,
                    "Author Link" : author_link,
                })

            print(f"  Found {len(quote_divs)} quotes on this page  [Total so far: {len(all_quotes)}]")

            # ── Find the 'Next' button for the next page ────
            next_btn = soup.find("li", class_="next")
            if next_btn and next_btn.find("a"):
                next_href = next_btn.find("a")["href"]
                url = BASE_URL + next_href
                page_num += 1
                time.sleep(1)                               # polite delay
            else:
                url = None                                  # no more pages

        except requests.exceptions.RequestException as e:
            print(f"  ERROR: Could not fetch page — {e}")
            break

    print(f"\n  SCRAPING COMPLETE!  Total quotes collected: {len(all_quotes)}")
    return all_quotes


# =============================================================
#  STEP 2 — CLEAN & ORGANIZE DATA
# =============================================================

def clean_data(df):
    """
    Cleans the scraped data:
    - Removes duplicates
    - Fixes formatting
    - Adds extra analysis columns
    """
    print("\n  Cleaning data...")

    original_count = len(df)
    df.drop_duplicates(subset=["Quote"], inplace=True)
    df.dropna(subset=["Quote", "Author"], inplace=True)
    df["Author"] = df["Author"].str.strip().str.title()
    df["Quote"]  = df["Quote"].str.strip()
    df.reset_index(drop=True, inplace=True)

    # Categorize quotes by length
    df["Quote Length"] = pd.cut(
        df["Word Count"],
        bins=[0, 10, 20, 30, 50, 999],
        labels=["Very Short", "Short", "Medium", "Long", "Very Long"]
    )

    removed = original_count - len(df)
    print(f"  Removed {removed} duplicates. Final count: {len(df)} quotes.")
    return df


# =============================================================
#  STEP 3 — SAVE FILES
# =============================================================

def save_files(df):
    """Save data to both CSV and Excel formats."""
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    print(f"\n  CSV saved  ->  {OUTPUT_CSV}")

    df.to_excel(OUTPUT_EXCEL, index=False, sheet_name="Quotes Data")
    print(f"  Excel saved  ->  {OUTPUT_EXCEL}")


# =============================================================
#  STEP 4 — CREATE 6 PROFESSIONAL CHARTS
# =============================================================

def make_charts(df):
    """
    Creates a professional 6-chart dashboard:
      1. Top 10 Most Quoted Authors
      2. Top 15 Most Popular Tags
      3. Quote Length Distribution (histogram)
      4. Quote Length Category Pie Chart
      5. Avg Word Count per Author (top 8)
      6. Tags per Quote Distribution
    """
    print("\n  Generating professional charts...")

    BLUE = ["#1E3A5F","#1B5299","#2563EB","#3B82F6",
            "#60A5FA","#93C5FD","#BFDBFE","#DBEAFE",
            "#172554","#1E40AF","#1D4ED8","#2563EB",
            "#3B82F6","#60A5FA","#93C5FD"]

    # Parse all individual tags
    all_tags = []
    df["Tags"].dropna().str.split(",").apply(
        lambda x: all_tags.extend([t.strip() for t in x if t.strip()])
    )
    tag_counts = pd.Series(Counter(all_tags)).sort_values(ascending=False)

    # ── Figure Setup ──────────────────────────────────────────
    fig = plt.figure(figsize=(20, 15), facecolor="#EEF2FF")
    fig.suptitle(
        "Quotes Data Analysis Dashboard\n"
        "Web Scraped from quotes.toscrape.com  |  CodeAlpha Internship — Task 1",
        fontsize=17, fontweight="bold", color="#1E3A5F", y=0.98
    )

    gs = gridspec.GridSpec(3, 2, figure=fig, hspace=0.55, wspace=0.35)

    # ── Chart 1: Top 10 Authors ───────────────────────────────
    ax1 = fig.add_subplot(gs[0, 0])
    top_authors = df["Author"].value_counts().head(10)
    bars = ax1.barh(top_authors.index[::-1], top_authors.values[::-1],
                    color=BLUE[:10], edgecolor="white", linewidth=0.8)
    for bar, val in zip(bars, top_authors.values[::-1]):
        ax1.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height()/2,
                 f"  {val} quotes", va="center", fontsize=8.5,
                 color="#1E3A5F", fontweight="bold")
    ax1.set_title("Top 10 Most Quoted Authors", fontsize=12,
                  fontweight="bold", color="#1E3A5F", pad=10)
    ax1.set_xlabel("Number of Quotes", fontsize=9, color="#374151")
    ax1.set_facecolor("#F8FAFF")
    ax1.tick_params(labelsize=8, colors="#374151")
    ax1.spines[["top","right"]].set_visible(False)
    ax1.set_xlim(0, top_authors.max() + 2)

    # ── Chart 2: Top 15 Tags ──────────────────────────────────
    ax2 = fig.add_subplot(gs[0, 1])
    top_tags = tag_counts.head(12)
    bars2 = ax2.bar(range(len(top_tags)), top_tags.values,
                    color=BLUE[:len(top_tags)], edgecolor="white", linewidth=0.8, width=0.7)
    ax2.set_xticks(range(len(top_tags)))
    ax2.set_xticklabels(top_tags.index, rotation=40, ha="right", fontsize=8)
    for bar, val in zip(bars2, top_tags.values):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                 str(val), ha="center", fontsize=8.5,
                 color="#1E3A5F", fontweight="bold")
    ax2.set_title("Top 12 Most Popular Tags", fontsize=12,
                  fontweight="bold", color="#1E3A5F", pad=10)
    ax2.set_ylabel("Frequency", fontsize=9, color="#374151")
    ax2.set_facecolor("#F8FAFF")
    ax2.tick_params(labelsize=8, colors="#374151")
    ax2.spines[["top","right"]].set_visible(False)

    # ── Chart 3: Word Count Distribution ─────────────────────
    ax3 = fig.add_subplot(gs[1, 0])
    ax3.hist(df["Word Count"], bins=18, color="#2563EB",
             edgecolor="white", linewidth=0.8, alpha=0.88)
    ax3.axvline(df["Word Count"].mean(), color="#1E3A5F",
                linestyle="--", linewidth=2,
                label=f"Mean: {df['Word Count'].mean():.1f} words")
    ax3.axvline(df["Word Count"].median(), color="#60A5FA",
                linestyle="-.", linewidth=2,
                label=f"Median: {df['Word Count'].median():.1f} words")
    ax3.set_title("Distribution of Quote Word Count", fontsize=12,
                  fontweight="bold", color="#1E3A5F", pad=10)
    ax3.set_xlabel("Word Count", fontsize=9, color="#374151")
    ax3.set_ylabel("Number of Quotes", fontsize=9, color="#374151")
    ax3.legend(fontsize=8, framealpha=0.6)
    ax3.set_facecolor("#F8FAFF")
    ax3.tick_params(labelsize=8)
    ax3.spines[["top","right"]].set_visible(False)

    # ── Chart 4: Quote Length Category Pie ───────────────────
    ax4 = fig.add_subplot(gs[1, 1])
    length_counts = df["Quote Length"].value_counts()
    pie_colors = ["#1E3A5F","#2563EB","#60A5FA","#BFDBFE","#EFF6FF"]
    wedges, texts, autotexts = ax4.pie(
        length_counts.values,
        labels=length_counts.index,
        autopct="%1.1f%%",
        colors=pie_colors[:len(length_counts)],
        startangle=140,
        pctdistance=0.75,
        wedgeprops=dict(edgecolor="white", linewidth=2.5)
    )
    for at in autotexts:
        at.set_fontsize(9); at.set_fontweight("bold"); at.set_color("white")
    for t in texts:
        t.set_fontsize(9); t.set_color("#1E3A5F")
    ax4.set_title("Quote Length Categories", fontsize=12,
                  fontweight="bold", color="#1E3A5F", pad=10)
    ax4.set_facecolor("#F8FAFF")

    # ── Chart 5: Avg Word Count per Author ────────────────────
    ax5 = fig.add_subplot(gs[2, 0])
    top_author_names = df["Author"].value_counts().head(8).index
    avg_words = (df[df["Author"].isin(top_author_names)]
                 .groupby("Author")["Word Count"]
                 .mean()
                 .sort_values(ascending=True))
    bars5 = ax5.barh(avg_words.index, avg_words.values,
                     color="#3B82F6", edgecolor="white", linewidth=0.8, alpha=0.9)
    for bar, val in zip(bars5, avg_words.values):
        ax5.text(bar.get_width() + 0.2, bar.get_y() + bar.get_height()/2,
                 f"  {val:.1f}", va="center", fontsize=8.5,
                 color="#1E3A5F", fontweight="bold")
    ax5.set_title("Avg Words per Quote (Top Authors)", fontsize=12,
                  fontweight="bold", color="#1E3A5F", pad=10)
    ax5.set_xlabel("Average Word Count", fontsize=9, color="#374151")
    ax5.set_facecolor("#F8FAFF")
    ax5.tick_params(labelsize=8)
    ax5.spines[["top","right"]].set_visible(False)
    ax5.set_xlim(0, avg_words.max() + 5)

    # ── Chart 6: Tags per Quote ───────────────────────────────
    ax6 = fig.add_subplot(gs[2, 1])
    tag_per_quote = df["Tag Count"].value_counts().sort_index()
    bars6 = ax6.bar(tag_per_quote.index.astype(str), tag_per_quote.values,
                    color=BLUE[:len(tag_per_quote)], edgecolor="white",
                    linewidth=0.8, width=0.7)
    for bar, val in zip(bars6, tag_per_quote.values):
        ax6.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
                 str(val), ha="center", fontsize=9,
                 color="#1E3A5F", fontweight="bold")
    ax6.set_title("How Many Tags per Quote?", fontsize=12,
                  fontweight="bold", color="#1E3A5F", pad=10)
    ax6.set_xlabel("Number of Tags", fontsize=9, color="#374151")
    ax6.set_ylabel("Number of Quotes", fontsize=9, color="#374151")
    ax6.set_facecolor("#F8FAFF")
    ax6.tick_params(labelsize=9)
    ax6.spines[["top","right"]].set_visible(False)

    # ── Save ──────────────────────────────────────────────────
    plt.savefig(OUTPUT_CHART, dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close()
    print(f"  Charts saved  ->  {OUTPUT_CHART}")


# =============================================================
#  STEP 5 — PRINT SUMMARY INSIGHTS
# =============================================================

def print_summary(df):
    """Prints key insights from the scraped data."""
    all_tags = []
    df["Tags"].dropna().str.split(",").apply(
        lambda x: all_tags.extend([t.strip() for t in x if t.strip()])
    )

    print("\n" + "="*55)
    print("  KEY INSIGHTS FROM YOUR SCRAPED DATA")
    print("="*55)
    print(f"  Total Quotes Scraped      : {len(df)}")
    print(f"  Unique Authors            : {df['Author'].nunique()}")
    print(f"  Unique Tags               : {len(set(all_tags))}")
    print(f"  Avg Words per Quote       : {df['Word Count'].mean():.1f}")
    print(f"  Longest Quote (words)     : {df['Word Count'].max()}")
    print(f"  Shortest Quote (words)    : {df['Word Count'].min()}")
    print(f"  Most Quoted Author        : {df['Author'].value_counts().index[0]}")
    print(f"  Most Popular Tag          : {Counter(all_tags).most_common(1)[0][0]}")
    print("="*55)


# =============================================================
#  MAIN — Run All Steps
# =============================================================

if __name__ == "__main__":
    print("\n" + "="*55)
    print("  CodeAlpha Internship | Task 1: Web Scraping")
    print("  Scraping: quotes.toscrape.com")
    print("="*55)

    # 1. Scrape
    raw_data = scrape_all_quotes()

    if not raw_data:
        print("\nERROR: No data scraped. Check internet connection.")
        raise SystemExit(1)

    # 2. Clean
    df = pd.DataFrame(raw_data)
    df = clean_data(df)

    # 3. Preview
    print(f"\n  PREVIEW (First 3 Quotes):\n")
    for i, row in df.head(3).iterrows():
        print(f"  [{i+1}] \"{row['Quote'][:70]}...\"")
        print(f"       — {row['Author']}  |  Tags: {row['Tags']}\n")

    # 4. Save CSV + Excel
    save_files(df)

    # 5. Charts
    make_charts(df)

    # 6. Summary
    print_summary(df)

    print("\n  YOUR PROJECT FILES:")
    print(f"    scraper.py          <- source code")
    print(f"    {OUTPUT_CSV}     <- dataset (CSV)")
    print(f"    {OUTPUT_EXCEL}  <- dataset (Excel)")
    print(f"    {OUTPUT_CHART} <- charts")
    print("\n  Upload all files to GitHub!")
    print("="*55 + "\n")
