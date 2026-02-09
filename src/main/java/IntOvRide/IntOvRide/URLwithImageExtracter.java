package IntOvRide.IntOvRide;

import com.microsoft.playwright.*;
import com.microsoft.playwright.options.LoadState;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;
import java.nio.file.*;
import java.util.List;

public class URLwithImageExtracter {

    private static final String HOME_URL =
            "https://www.vogue.in/web-stories/";

    private static final String OUTPUT_DIR =
            "D:\\Transformation Project\\Interactive\\Cover_Image";

    private static final String EXCEL_PATH =
            OUTPUT_DIR + "\\WebStories.xlsx";

    public static void main(String[] args) {

        try (Playwright playwright = Playwright.create()) {

            Browser browser = playwright.chromium().launch(
                    new BrowserType.LaunchOptions()
                            .setHeadless(false)
                            .setChannel("chrome")
            );

            BrowserContext context = browser.newContext();
            Page page = context.newPage();

            // Global timeout
            page.setDefaultTimeout(30000);

            // ✅ FIX: Safe navigation for ad-heavy site
            page.navigate(
                    HOME_URL,
                    new Page.NavigateOptions()
  //                         .setWaitUntil(LoadState.DOMCONTENTLOADED)
                            .setTimeout(60000)
            );

            // ✅ Wait ONLY for required elements
            page.waitForSelector(
                    "div.col.two-cols, div.col.three-cols",
                    new Page.WaitForSelectorOptions().setTimeout(20000)
            );

            Files.createDirectories(Paths.get(OUTPUT_DIR));

            // Create Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("WebStories");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Story URL");
            header.createCell(1).setCellValue("Cover Image Path");
            header.createCell(2).setCellValue("Title");

            // Extract cards
            List<ElementHandle> cards =
                    page.querySelectorAll("div.col.two-cols, div.col.three-cols");

            int rowIndex = 1;
            int imageIndex = 1;

            for (ElementHandle card : cards) {

                ElementHandle linkEl = card.querySelector("a.product_image");
                if (linkEl == null) continue;

                String storyUrl = linkEl.getAttribute("href");
                if (storyUrl == null || storyUrl.isEmpty()) continue;

                ElementHandle imgEl = linkEl.querySelector("img");
                if (imgEl == null) continue;

                String imgUrl = imgEl.getAttribute("src");
                if (imgUrl == null || imgUrl.isEmpty()) continue;

                ElementHandle titleEl = card.querySelector("a.product-title");
                String titleText = titleEl != null
                        ? titleEl.innerText().trim()
                        : "TITLE_NOT_FOUND";

                // Download image
                String extension =
                        imgUrl.substring(imgUrl.lastIndexOf(".")).split("\\?")[0];

                String imageName = "cover_" + imageIndex + extension;
                Path imagePath = Paths.get(OUTPUT_DIR + "\\" + imageName);

                try (InputStream in = new URL(imgUrl).openStream()) {
                    Files.copy(in, imagePath, StandardCopyOption.REPLACE_EXISTING);
                }

                // Write Excel
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(storyUrl);
                row.createCell(1).setCellValue(imagePath.toString());
                row.createCell(2).setCellValue(titleText);

                System.out.println("✔ Extracted:");
                System.out.println("  URL   : " + storyUrl);
                System.out.println("  Title : " + titleText);

                imageIndex++;
            }

            // Save Excel
            try (FileOutputStream fos = new FileOutputStream(EXCEL_PATH)) {
                workbook.write(fos);
            }
            workbook.close();

            browser.close();

            System.out.println("=================================");
            System.out.println("Extraction Completed Successfully");
            System.out.println("Excel saved at: " + EXCEL_PATH);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
