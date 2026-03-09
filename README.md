\# CHMeetings Check Import



Extract contribution data from scanned check images in a PDF and produce an Excel file ready to import into \[CHMeetings](https://www.chmeetings.com/) under \*\*Contributions → Import\*\*.



All processing runs locally on your machine — nothing is uploaded to the cloud.



\## Features



\- \*\*OCR extraction\*\* of name, amount, and check number from check images

\- \*\*Interactive review\*\* with check image displayed in a popup window

\- \*\*Contact matching\*\* with fuzzy name lookup against a CHMeetings people export

\- \*\*Autocomplete\*\* when typing names during edit mode (Tab to accept)

\- \*\*Navigation\*\* — go back to previous checks to correct mistakes

\- \*\*Running totals\*\* after each check and a summary at the end



\## Requirements



\- Python 3.10+

\- Dependencies: `pip install pymupdf easyocr openpyxl Pillow`



First run downloads the EasyOCR English model (~100 MB).



\## Usage



```

python check\_to\_chmeetings.py "checks.pdf" --review

python check\_to\_chmeetings.py "checks.pdf" --review --contacts "people.xlsx"

python check\_to\_chmeetings.py "checks.pdf" --review --contacts "people.xlsx" --fund "Tithes" --date "03/08/2026"

python check\_to\_chmeetings.py "checks.pdf" --review --batch "March 2026" --deposit-date "03/08/2026"

```



\### Review mode keys



| Key | Action |

|-----|--------|

| \*\*A\*\* | Accept the entry as shown |

| \*\*M\*\* | Use the matched contact name (when a match is found) |

| \*\*E\*\* | Edit fields (Tab autocompletes names from contacts) |

| \*\*S\*\* | Skip this check |

| \*\*P\*\* | Go back to the previous check |



\### Options



| Flag | Description |

|------|-------------|

| `--review` | Interactive review with image display |

| `--contacts FILE` | CHMeetings people export xlsx for name matching |

| `--fund NAME` | Default fund for all entries |

| `--date MM/DD/YYYY` | Contribution date for all entries |

| `--batch NAME` | Batch name |

| `--batch-number NUM` | Batch number |

| `--deposit-date MM/DD/YYYY` | Deposit date |

| `--payment-method METHOD` | Payment method (default: Check) |

| `--verbose` | Show OCR text by region for debugging |



\## License



MIT

