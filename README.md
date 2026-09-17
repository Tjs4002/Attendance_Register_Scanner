# Attendance register scanner

Run `node server.js` and open http://localhost:3000. Configure your Gemini API key in Settings.

Use the **Theme** selector in the header to choose **Light**, **Dark**, or **System**. Your choice is saved in this browser; System follows your device's appearance settings. Register images keep their original colors in every theme.

The interface follows three steps:

1. **Upload:** use **Choose files**, **Take photo**, or drag and drop. Check the page previews, remove unwanted pages, then click **Scan register**.
2. **Review:** compare the original register and student table, edit cells, and use **Class for this page** with **Apply to Page** for a page-wide correction. **More options** contains name formatting, roll numbering, default section, bulk class changes, birthdate generation, and copying the table.
3. **Download:** check the student count and choose Excel or CSV. Downloads include the full list regardless of search or class filters. You can return to Review without losing edits.

Each PDF page appears separately in the page selector. Scanned and printed Marathi or English registers use the same vision OCR and name transliteration workflow. PDF rendering loads PDF.js from a version-pinned CDN and requires internet access. PDF limits are 50 MB and 100 pages per file; password-protected documents must be unlocked before uploading.

**Add pages** preserves existing student edits. Return to **Upload** and click **Scan register** to retry unfinished pages without duplicating completed pages. If scanning opens Settings because no API key is configured, saving a key continues that requested scan. Saving Settings otherwise does not start scanning. The header menu contains **Start over**.

Class detection recognizes Marathi digits and headings. Camera filenames and dates are not interpreted as classes. Explicit filenames such as `register_class6.jpg` or `std_6.pdf` provide a fallback hint; readable class headings take priority so mixed-class pages remain supported. Unknown classes remain blank for review.

Run regression checks with `node --test tests/scanner.test.cjs` (or `npm test` where npm is installed). Tests cover class parsing, page rendering orchestration, cancellation, password failures, and incremental scan retries using mocked PDF.js and Gemini responses. Real PDF rendering and recognition accuracy should also be checked with representative registers and a configured Gemini key.
