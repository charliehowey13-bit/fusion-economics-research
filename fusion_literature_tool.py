import requests
import pandas as pd
import time
from datetime import datetime

# CONFIGURATION

EMAIL = "charliehowey13@gmail.com" 

KEYWORDS = [
    "tokamak economics",
    "stellarator economics",
    "inertial confinement fusion economics",
    "magnetized target fusion economics",
    "fusion commercialization",
    "fusion technology readiness level",
    "fusion cost analysis",
    "levelized cost of electricity fusion",
    "fusion reactor capital cost",
    "tritium breeding economics",
    "fusion energy systems analysis",
    "multi criteria decision analysis fusion"
]

BASE_URL = "https://api.openalex.org/works"

# SEARCH OPENALEX

def search_openalex(keyword, limit=50):

    print(f"Searching: {keyword}")

    params = {
        "search": keyword,
        "per_page": limit,
        "select": ",".join([
            "id",
            "title",
            "publication_year",
            "doi",
            "cited_by_count",
            "authorships",
            "primary_location"
        ])
    }

    if EMAIL:
        params["mailto"] = EMAIL

    try:
        response = requests.get(BASE_URL, params=params)

        if response.status_code != 200:
            print(f"Error {response.status_code}")
            return []

        return response.json().get("results", [])

    except Exception as e:
        print(e)
        return []


# CATEGORY DETECTION

def categorize(title):

    title = (title or "").lower()

    if "tokamak" in title:
        return "Tokamak"

    if "stellarator" in title:
        return "Stellarator"

    if "inertial confinement" in title:
        return "ICF"

    if "magnetized target" in title:
        return "MTF"

    if "tritium" in title:
        return "Tritium"

    if "econom" in title:
        return "Economics"

    return "General Fusion"


# ---------------------------
# CITATION SCORE
# ---------------------------

def score_paper(citations, year):

    current_year = datetime.now().year

    age = max(1, current_year - (year or current_year))

    citation_score = citations

    recency_bonus = max(0, 10 - age)

    return citation_score + recency_bonus


# MAIN COLLECTION

all_papers = []

seen_ids = set()

for keyword in KEYWORDS:

    results = search_openalex(keyword)

    for paper in results:

        paper_id = paper.get("id")

        if paper_id in seen_ids:
            continue

        seen_ids.add(paper_id)

        authors = []

        for author in paper.get("authorships", []):
            try:
                authors.append(
                    author["author"]["display_name"]
                )
            except:
                pass

        title = paper.get("title")

        year = paper.get("publication_year")

        citations = paper.get("cited_by_count", 0)

        all_papers.append({

            "Category":
                categorize(title),

            "Title":
                title,

            "Year":
                year,

            "Authors":
                ", ".join(authors),

            "Citations":
                citations,

            "DOI":
                paper.get("doi"),

            "Open Access URL":
                paper.get("primary_location", {})
                     .get("landing_page_url"),

            "Search Keyword":
                keyword,

            "Score":
                score_paper(citations, year)

        })

    time.sleep(0.25)

# DATAFRAMES

df = pd.DataFrame(all_papers)

df = df.sort_values(
    by="Score",
    ascending=False
)

top_papers = df.head(100)

# EXPORT EXCEL

output_file = "fusion_literature_review.xlsx"

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="All Papers",
        index=False
    )

    top_papers.to_excel(
        writer,
        sheet_name="Top 100",
        index=False
    )

    for category in df["Category"].unique():

        subset = df[
            df["Category"] == category
        ]

        subset.to_excel(
            writer,
            sheet_name=category[:31],
            index=False
        )

print()
print("=" * 50)
print("DONE")
print(f"Total papers: {len(df)}")
print(f"Saved to: {output_file}")
print("=" * 50)
