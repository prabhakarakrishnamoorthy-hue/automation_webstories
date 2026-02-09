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

public class SliderToInteractiveHTML1 {

    // ===== CONFIG =====
    private static final String EXCEL_PATH =
            "D:\\Transformation Project\\Interactive\\Input\\Input_URL.xlsx";

    private static final String OUTPUT_DIR =
            "D:\\Transformation Project\\Interactive\\Output";

    private static final String COVER_IMAGE_DIR =
            "D:\\Transformation Project\\Interactive\\Cover_Image";

    public static void main(String[] args) throws Exception {

        FileInputStream fis = new FileInputStream(EXCEL_PATH);
        Workbook workbook = WorkbookFactory.create(fis);
        Sheet sheet = workbook.getSheetAt(0);

        Files.createDirectories(Paths.get(COVER_IMAGE_DIR));

        try (Playwright playwright = Playwright.create()) {

            Browser browser = playwright.chromium().launch(
                    new BrowserType.LaunchOptions()
                            .setHeadless(false)
                            .setChannel("chrome")
            );

            BrowserContext context = browser.newContext();
            Page page = context.newPage();

            int index = 1;

            for (int r = 0; r <= sheet.getLastRowNum(); r++) {

                Row row = sheet.getRow(r);
                if (row == null) continue;

                Cell urlCell = row.getCell(0);
                if (urlCell == null || urlCell.getCellType() != CellType.STRING) continue;

                String url = urlCell.getStringCellValue().trim();
                if (url.isEmpty()) continue;

                System.out.println("====================================");
                System.out.println("Processing URL: " + url);

                String pageFolder = OUTPUT_DIR + "\\interactive_" + index;
                String assetFolder = pageFolder + "\\assets";

                Files.createDirectories(Paths.get(assetFolder + "\\css"));
                Files.createDirectories(Paths.get(assetFolder + "\\js"));
                Files.createDirectories(Paths.get(assetFolder + "\\images"));

                // 1️⃣ Load page
                page.navigate(url);
                page.waitForLoadState(LoadState.NETWORKIDLE);
                page.waitForTimeout(3000);

                // 2️⃣ EXTRACT COVER IMAGE USING PLAYWRIGHT (WEB STORY SAFE)
                Object coverImageUrl = page.evaluate(
                        "() => {" +
                        "  const ampImg = document.querySelector('amp-story-page amp-img img');" +
                        "  if (ampImg && ampImg.src) return ampImg.src;" +
                        "  const bg = document.querySelector('amp-story-page [style*=\"background-image\"]');" +
                        "  if (bg) {" +
                        "    const match = bg.style.backgroundImage.match(/url\\([\"']?(.*?)[\"']?\\)/);" +
                        "    if (match) return match[1];" +
                        "  }" +
                        "  const meta = document.querySelector('meta[property=\"og:image\"]');" +
                        "  return meta ? meta.content : null;" +
                        "}"
                );

                if (coverImageUrl != null && !((String) coverImageUrl).isEmpty()) {
                    saveCoverImageFromUrl((String) coverImageUrl, index);
                } else {
                    System.out.println("⚠️ Cover image not found via DOM.");
                }

                // 3️⃣ Capture FINAL DOM
                String html = page.content();
                Document doc = Jsoup.parse(html, url);

                // 4️⃣ Extract Title → Excel Column B
                String title = extractTitle(doc);
                Cell titleCell = row.getCell(1);
                if (titleCell == null) titleCell = row.createCell(1);
                titleCell.setCellValue(title);

                System.out.println("Title Captured: " + title);

                // 5️⃣ Download assets
                downloadAssets(doc, assetFolder);

                // 6️⃣ Save interactive.html
                Files.write(
                        Paths.get(pageFolder + "\\interactive.html"),
                        doc.outerHtml().getBytes()
                );

                System.out.println("Interactive HTML saved.");

                index++;
            }

            browser.close();
        }

        fis.close();

        FileOutputStream fos = new FileOutputStream(EXCEL_PATH);
        workbook.write(fos);
        fos.close();
        workbook.close();

        System.out.println("====================================");
        System.out.println("PROCESS COMPLETED SUCCESSFULLY");
    }

    // ================= COVER IMAGE SAVE =================
    private static void saveCoverImageFromUrl(String imageUrl, int index) {
        try {
            String extension = imageUrl.contains(".")
                    ? imageUrl.substring(imageUrl.lastIndexOf(".")).split("\\?")[0]
                    : ".jpg";

            String fileName = "cover_" + index + extension;
            Path targetPath = Paths.get(COVER_IMAGE_DIR + "\\" + fileName);

            try (InputStream in = new URL(imageUrl).openStream()) {
                Files.copy(in, targetPath, StandardCopyOption.REPLACE_EXISTING);
            }

            System.out.println("✅ Cover Image Saved: " + fileName);

        } catch (Exception e) {
            System.out.println("❌ Cover image download failed: " + e.getMessage());
        }
    }

    // ================= TITLE EXTRACTION =================
    private static String extractTitle(Document doc) {
        Element title =
                doc.selectFirst("h1") != null ? doc.selectFirst("h1") :
                doc.selectFirst("h2") != null ? doc.selectFirst("h2") :
                doc.selectFirst("meta[property=og:title]");

        return title != null
                ? (title.tagName().equals("meta") ? title.attr("content") : title.text())
                : "TITLE_NOT_FOUND";
    }

    // ================= ASSET DOWNLOADER =================
    private static void downloadAssets(Document doc, String assetBase) {

        for (Element img : doc.select("img[src]")) {
            downloadAndReplace(img, "src", assetBase + "\\images");
        }

        for (Element css : doc.select("link[rel=stylesheet][href]")) {
            downloadAndReplace(css, "href", assetBase + "\\css");
        }

        for (Element js : doc.select("script[src]")) {
            downloadAndReplace(js, "src", assetBase + "\\js");
        }
    }

    private static void downloadAndReplace(Element element, String attr, String targetDir) {
        try {
            String absUrl = element.absUrl(attr);
            if (absUrl.isEmpty()) return;

            String fileName = absUrl.substring(absUrl.lastIndexOf("/") + 1);
            if (fileName.contains("?")) {
                fileName = fileName.substring(0, fileName.indexOf("?"));
            }

            Path targetPath = Paths.get(targetDir + "\\" + fileName);

            try (InputStream in = new URL(absUrl).openStream()) {
                Files.copy(in, targetPath, StandardCopyOption.REPLACE_EXISTING);
            }

            String folderName = Paths.get(targetDir).getFileName().toString();
            element.attr(attr, "assets/" + folderName + "/" + fileName);

        } catch (Exception e) {
            System.out.println("Skipping asset: " + e.getMessage());
        }
    }
}
