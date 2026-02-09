package IntOvRide.IntOvRide;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;

import org.apache.poi.ss.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.*;
import java.net.URL;
import java.nio.file.*;
import java.util.*;

public class SliderToInteractiveHTML {

    // ===== CONFIG =====
    private static final String EXCEL_PATH =
            "D:\\Transformation Project\\Interactive\\Input\\Input_URL.xlsx";

    private static final String OUTPUT_DIR =
            "D:\\Transformation Project\\Interactive\\Output";

    public static void main(String[] args) throws Exception {

        List<String> urls = readUrlsFromExcel(EXCEL_PATH);

        try (Playwright playwright = Playwright.create()) {

            // ✅ Use local Chrome → NO browser download
            Browser browser = playwright.chromium().launch(
                    new BrowserType.LaunchOptions()
                            .setHeadless(false)
                            .setChannel("chrome")
            );

            BrowserContext context = browser.newContext();
            Page page = context.newPage();

            int count = 1;

            for (String url : urls) {

                System.out.println("=====================================");
                System.out.println("Processing URL: " + url);

                String pageFolder = OUTPUT_DIR + "\\interactive_" + count;
                String assetFolder = pageFolder + "\\assets";

                Files.createDirectories(Paths.get(assetFolder + "\\css"));
                Files.createDirectories(Paths.get(assetFolder + "\\js"));
                Files.createDirectories(Paths.get(assetFolder + "\\images"));

                // 1️⃣ Load URL
                page.navigate(url);
                page.waitForLoadState(LoadState.DOMCONTENTLOADED);

                // 2️⃣ Wait for legacy hero + title + slider
                page.waitForSelector(
                        "img, h1, h2, .hero, .banner, .title, .slider, .carousel, .swiper",
                        new Page.WaitForSelectorOptions().setTimeout(20000)
                );

                // Extra safety wait for JS-rendered sliders
                page.waitForTimeout(3000);

                // 3️⃣ Capture FINAL rendered DOM
                String htmlContent = page.content();

                // 4️⃣ Parse HTML
                Document document = Jsoup.parse(htmlContent, url);

                // 5️⃣ Download Assets
                downloadAssets(document, assetFolder);

                // 6️⃣ Save Interactive HTML
                Files.write(
                        Paths.get(pageFolder + "\\interactive.html"),
                        document.outerHtml().getBytes()
                );

                System.out.println("Saved interactive HTML at:");
                System.out.println(pageFolder + "\\interactive.html");

                count++;
            }

            browser.close();
        }
    }

    // ================= EXCEL READER =================
    private static List<String> readUrlsFromExcel(String excelPath) throws Exception {
        List<String> urls = new ArrayList<>();

        FileInputStream fis = new FileInputStream(excelPath);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            Cell cell = row.getCell(0);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String value = cell.getStringCellValue().trim();
                if (!value.isEmpty()) {
                    urls.add(value);
                }
            }
        }
        workbook.close();
        return urls;
    }

    // ================= ASSET DOWNLOADER =================
    private static void downloadAssets(Document doc, String assetBase) {

        // Images (hero + slider + thumbnails)
        for (Element img : doc.select("img[src]")) {
            downloadAndReplace(img, "src", assetBase + "\\images");
        }

        // CSS
        for (Element css : doc.select("link[rel=stylesheet][href]")) {
            downloadAndReplace(css, "href", assetBase + "\\css");
        }

        // JavaScript
        for (Element js : doc.select("script[src]")) {
            downloadAndReplace(js, "src", assetBase + "\\js");
        }
    }

    private static void downloadAndReplace(Element element, String attr, String targetDir) {
        try {
            String absUrl = element.absUrl(attr);
            if (absUrl == null || absUrl.isEmpty()) return;

            String fileName = absUrl.substring(absUrl.lastIndexOf("/") + 1);
            if (fileName.contains("?")) {
                fileName = fileName.substring(0, fileName.indexOf("?"));
            }

            Path targetPath = Paths.get(targetDir + "\\" + fileName);

            try (InputStream in = new URL(absUrl).openStream()) {
                Files.copy(in, targetPath, StandardCopyOption.REPLACE_EXISTING);
            }

            String folderName = Paths.get(targetDir).getFileName().toString();
            String relativePath = "assets/" + folderName + "/" + fileName;

            element.attr(attr, relativePath);

        } catch (Exception e) {
            System.out.println("Skipping asset: " + e.getMessage());
        }
    }
}
