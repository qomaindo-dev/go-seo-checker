package main

import (
	"context"
	"fmt"
	"net/http"
	"os"
	"regexp"
	"runtime"
	"strings"
	"sync"
	"time"

	"github.com/PuerkitoBio/goquery"
	"github.com/xuri/excelize/v2"
)

type Job struct {
	Row int
	URL string
}

type Result struct {
	Row   int
	URL   string
	Text  string
	OK    bool
	Error error
}

func normalizeRobotsValue(s string) string {
	s = strings.ToLower(strings.TrimSpace(s))
	// normalize spaces around comma
	re := regexp.MustCompile(`\s*,\s*`)
	s = re.ReplaceAllString(s, ",")
	return s
}

func containsNoindexOrNofollow(value string) bool {
	v := normalizeRobotsValue(value)
	return strings.Contains(v, "noindex") || strings.Contains(v, "nofollow")
}

func checkURL(ctx context.Context, client *http.Client, url string) (lines []string, ok bool, err error) {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, url, nil)
	if err != nil {
		return nil, false, err
	}
	req.Header.Set("User-Agent", "Mozilla/5.0 (compatible; SEOChecker/1.0; +https://example.com)")

	resp, err := client.Do(req)
	if err != nil {
		return nil, false, err
	}
	defer resp.Body.Close()

	found := false

	// To Check X-Robots-Tag header (can appear multiple times, excelize reads with Get, but net/http merges)
	xRobots := resp.Header.Values("X-Robots-Tag")
	if len(xRobots) == 0 {
		if v := resp.Header.Get("X-Robots-Tag"); v != "" {
			xRobots = []string{v}
		}
	}
	for _, v := range xRobots {
		if strings.TrimSpace(v) == "" {
			continue
		}
		if containsNoindexOrNofollow(v) {
			found = true
			lines = append(lines, fmt.Sprintf("❌ X-Robots-Tag found: %s", normalizeRobotsValue(v)))
		}
	}

	// To Parse HTML meta robots/googlebot
	doc, err := goquery.NewDocumentFromReader(resp.Body)
	if err != nil {
		// if the HTML is broken, still return what was found from the header
		if found {
			return lines, false, nil
		}
		return nil, false, err
	}

	doc.Find("meta").Each(func(i int, s *goquery.Selection) {
		name, _ := s.Attr("name")
		content, _ := s.Attr("content")
		name = strings.ToLower(strings.TrimSpace(name))
		content = normalizeRobotsValue(content)

		if (name == "robots" || name == "googlebot") && containsNoindexOrNofollow(content) {
			found = true
			lines = append(lines, fmt.Sprintf("❌ Meta %s found: %s", name, content))
		}
	})

	if !found {
		return []string{"✅ Noindex / nofollow was found on the link"}, true, nil
	}
	return lines, false, nil
}

func main() {
	// ====== INPUT / OUTPUT ======
	// Default according to your file. Can be changed via arg:
	// go run . input.xlsx output.xlsx
	inFile := "List-Link.xlsx"
	outFile := "Link-List_RESULT.xlsx"
	if len(os.Args) >= 2 {
		inFile = os.Args[1]
	}
	if len(os.Args) >= 3 {
		outFile = os.Args[2]
	}

	// ====== OPEN EXCEL ======
	f, err := excelize.OpenFile(inFile)
	if err != nil {
		fmt.Println("Gagal buka file:", err)
		os.Exit(1)
	}
	defer func() { _ = f.Close() }()

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		fmt.Println("Sheet tidak ditemukan di Excel.")
		os.Exit(1)
	}
	sheet := sheets[0] // use the first sheet

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println("Gagal baca rows:", err)
		os.Exit(1)
	}
	if len(rows) < 2 {
		fmt.Println("Data kosong. Pastikan ada header + minimal 1 link.")
		os.Exit(1)
	}

	// Find the "Link" and "Result" columns from the header row.
	linkCol := -1
	hasilCol := -1
	for i, h := range rows[0] {
		hh := strings.TrimSpace(strings.ToLower(h))
		if hh == "link" {
			linkCol = i
		}
		if hh == "hasil" {
			hasilCol = i
		}
	}
	if linkCol == -1 {
		fmt.Println(`Kolom header "Link" tidak ditemukan.`)
		os.Exit(1)
	}
	if hasilCol == -1 {
		// If it doesn't exist, create a new column after Link
		hasilCol = linkCol + 1
		hasilCell, _ := excelize.CoordinatesToCellName(hasilCol+1, 1)
		_ = f.SetCellValue(sheet, hasilCell, "Hasil")
	}

	// ====== HTTP CLIENT ======
	client := &http.Client{
		Timeout: 20 * time.Second,
		// Following default redirect (net/http) — it's OK
	}

	// ====== JOBS (get all URLs from excel) ======
	var jobs []Job
	for r := 2; r <= len(rows); r++ { // start row 2 (After header)
		linkCell, _ := excelize.CoordinatesToCellName(linkCol+1, r)
		url, _ := f.GetCellValue(sheet, linkCell)
		url = strings.TrimSpace(url)
		if url == "" {
			continue
		}
		jobs = append(jobs, Job{Row: r, URL: url})
	}

	// ====== CONCURRENCY WORKER POOL ======
	workerCount := runtime.NumCPU()
	if workerCount < 4 {
		workerCount = 4
	}
	jobCh := make(chan Job)
	resCh := make(chan Result)

	var wg sync.WaitGroup
	for i := 0; i < workerCount; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for job := range jobCh {
				ctx, cancel := context.WithTimeout(context.Background(), 25*time.Second)
				lines, ok, err := checkURL(ctx, client, job.URL)
				cancel()

				if err != nil {
					resCh <- Result{
						Row:   job.Row,
						URL:   job.URL,
						Text:  "❌ Error: " + err.Error(),
						OK:    false,
						Error: err,
					}
					continue
				}
				resCh <- Result{
					Row:  job.Row,
					URL:  job.URL,
					Text: strings.Join(lines, "\n"),
					OK:   ok,
				}
			}
		}()
	}

	go func() {
		for _, j := range jobs {
			jobCh <- j
		}
		close(jobCh)
		wg.Wait()
		close(resCh)
	}()

	// ====== STYLE: wrap text in the results column ======
	wrapStyle, _ := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "top",
		},
	})

	greenStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "006100"}, // dark green
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "top",
		},
	})
	redStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Color: "9C0006"}, // dark red
		Alignment: &excelize.Alignment{
			WrapText: true,
			Vertical: "top",
		},
	})

	// Set column width for easy reading
	linkColName, _ := excelize.ColumnNumberToName(linkCol + 1)
	hasilColName, _ := excelize.ColumnNumberToName(hasilCol + 1)
	_ = f.SetColWidth(sheet, linkColName, linkColName, 50)
	_ = f.SetColWidth(sheet, hasilColName, hasilColName, 60)

	// ====== WRITE RESULTS ======
	for res := range resCh {
		hasilCell, _ := excelize.CoordinatesToCellName(hasilCol+1, res.Row)
		_ = f.SetCellValue(sheet, hasilCell, res.Text)

		// style
		if strings.HasPrefix(res.Text, "✅") {
			_ = f.SetCellStyle(sheet, hasilCell, hasilCell, greenStyle)
		} else if strings.HasPrefix(res.Text, "❌") {
			_ = f.SetCellStyle(sheet, hasilCell, hasilCell, redStyle)
		} else {
			_ = f.SetCellStyle(sheet, hasilCell, hasilCell, wrapStyle)
		}

		// Auto row height is more convenient (makes multiline rows legible)
		// Excelize doesn't have a pure "auto height," but you can set it a bit higher:
		_ = f.SetRowHeight(sheet, res.Row, 45)
	}

	// Header style
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Bold: true},
		Alignment: &excelize.Alignment{Horizontal: "center", Vertical: "center"},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	linkHeader, _ := excelize.CoordinatesToCellName(linkCol+1, 1)
	hasilHeader, _ := excelize.CoordinatesToCellName(hasilCol+1, 1)
	_ = f.SetCellStyle(sheet, linkHeader, linkHeader, headerStyle)
	_ = f.SetCellStyle(sheet, hasilHeader, hasilHeader, headerStyle)
	_ = f.SetRowHeight(sheet, 1, 22)

	// Save output
	if err := f.SaveAs(outFile); err != nil {
		fmt.Println("Failed to save output:", err)
		os.Exit(1)
	}

	fmt.Println("Finish. Output:", outFile)
}
