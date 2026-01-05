
# go-seo-checker

An automation test created using native Golang to perform SEO audits on websites that aims to detect whether there are "noindex, nofollow" elements on a website page based on the URL provided in an Excel file (multi-URL).

This program was created based on a case study I encountered in my freelance work. I hope this simple automation tool will be useful for all developers.

Thank you, and don't forget to follow me.


## How to run it?

To run it, there are 2 ways below:

### 1. Put your Excel file in the same folder, then:

```bash
  go run . Link-List.xlsx Link-List_RESULT.xlsx
```

### 2. Without arguments is also possible (use default file name):

```bash
  go run .
```
