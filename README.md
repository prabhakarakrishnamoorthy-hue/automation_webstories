# Automation Web Stories (IntOvRide)

Java automation to capture **web stories** (e.g. AMP/Google Web Stories), download their assets, and export them as offline interactive HTML. Also supports scraping story listings and cover images from index pages.

## Features

- **URL → Interactive HTML** – Read URLs from Excel, open each in Chrome via Playwright, capture the rendered DOM, download images/CSS/JS, and save self-contained `interactive.html` per URL.
- **Cover image & title extraction** – Extract cover images (AMP story or `og:image`) and page titles, save covers to a folder and write titles back to Excel.
- **Story listing scraper** – Visit a web-stories index URL, scrape story cards (URL, cover image, title), download cover images, and write results to an Excel file.

## Tech Stack

- **Java** (Maven)
- [Playwright for Java](https://playwright.dev/java/) – browser automation (Chrome)
- [Apache POI](https://poi.apache.org/) – read/write Excel (`.xlsx`)
- [Jsoup](https://jsoup.org/) – HTML parsing and asset URL resolution

## Prerequisites

- **JDK 8+**
- **Maven 3.x**
- **Chrome** installed (Playwright is configured to use `channel: "chrome"`)

## Build

```bash
mvn clean compile
```

## Run

Each entry point is a `main` in a different class. Run the one you need:

```bash
# Capture pages from Excel URLs → interactive HTML + assets
mvn exec:java -Dexec.mainClass="IntOvRide.IntOvRide.SliderToInteractiveHTML"

# Same + cover image extraction + title written to Excel
mvn exec:java -Dexec.mainClass="IntOvRide.IntOvRide.SliderToInteractiveHTML1"

# Scrape web-stories index page → Excel + cover images
mvn exec:java -Dexec.mainClass="IntOvRide.IntOvRide.URLwithImageExtracter"
```

Or run the corresponding class from your IDE (e.g. Run → SliderToInteractiveHTML, SliderToInteractiveHTML1, or URLwithImageExtracter).

## Configuration

Paths are defined as constants at the top of each class. Update them for your environment:

| Class | What to set |
|-------|-------------|
| **SliderToInteractiveHTML** | `EXCEL_PATH` (input URLs), `OUTPUT_DIR` (interactive HTML + assets) |
| **SliderToInteractiveHTML1** | Same as above + `COVER_IMAGE_DIR` (where cover images are saved) |
| **URLwithImageExtracter** | `HOME_URL` (web-stories index), `OUTPUT_DIR` (covers + Excel path); Excel path is `OUTPUT_DIR + "\\WebStories.xlsx"` |

Default values use Windows paths (e.g. `D:\Transformation Project\Interactive\...`). Change to your desired directories (use `\\` on Windows or `/` and adjust code if you move to Unix).

## Project Structure

```
src/main/java/IntOvRide/IntOvRide/
├── SliderToInteractiveHTML.java   # URLs from Excel → interactive HTML + assets
├── SliderToInteractiveHTML1.java  # Same + cover images + title in Excel
└── URLwithImageExtracter.java     # Scrape index page → Excel + cover images
```

- **SliderToInteractiveHTML** – Reads URLs from column A of an Excel file, visits each URL, waits for hero/slider-like content, captures DOM, downloads assets into `assets/images`, `assets/css`, `assets/js`, and saves `interactive.html` in `interactive_1`, `interactive_2`, …
- **SliderToInteractiveHTML1** – Same flow plus: extracts cover image (AMP story or `og:image`), saves to `COVER_IMAGE_DIR`; extracts title (h1/h2 or og:title) and writes it to column B in the same Excel file.
- **URLwithImageExtracter** – Navigates to a fixed index URL (e.g. Vogue web-stories), finds story cards (`div.col.two-cols`, `div.col.three-cols`), extracts story URL, cover image URL, and title, downloads cover images, and writes one row per story to `WebStories.xlsx` (Story URL, Cover Image Path, Title).

## Excel Format

- **Input (SliderToInteractiveHTML / SliderToInteractiveHTML1)**  
  - Column A: story/page URLs (one per row).  
  - SliderToInteractiveHTML1 also writes the extracted title to column B.

- **Output (URLwithImageExtracter)**  
  - `WebStories.xlsx`: columns **Story URL**, **Cover Image Path**, **Title**.

## License

See project or repository for license information.
