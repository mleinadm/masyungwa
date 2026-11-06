# Masyungwa Comprehensive School – Class Register

Pupil records for Masyungwa Comprehensive School, organized by class year.

## Files

| File | Description |
|------|-------------|
| `classes.md` | Source markdown with all pupil lists per class year |
| `classes.xlsx` | Generated Excel workbook (two sheets: **All Classes** and **Missing Classes**) |
| `create_excel.py` | Python script that parses `classes.md` → `classes.xlsx` |

## Excel Workbook

### Sheet 1: All Classes
- **Columns:** Name, Class Year, Champion
- Lists every pupil with their class year and whether they are a champion
- 318 students across 25 classes (as of latest update)

### Sheet 2: Missing Classes
- **Columns:** Year, Status
- Years between **1975–2019** with no full pupil list on record
- 20 missing years: 1976, 1977, 1980–1984, 1986, 1988, 1991, 1992, 1995, 2011–2014, 2016–2019

## Champion Pupils

| Year | Champion(s) |
|------|-------------|
| 1987 | Peter Musili |
| 1989 | Dickson Mwali |
| 1990 | Muo Kasina |
| 1996 | Sammy Kitheka |
| 1997 | Kithome Mutua / John Malekwa |
| 1998 | Richard Muatha / Sammy Maithya / Ivuti Muatha |
| 1999 | Nzou Maluki (CEO) |
| 2000 | Mwangangi Kilonzo |
| 2001 | Muthugi Maluki |
| 2002 | Maluki Muthoka |
| 2003 | Mwasya Muthoni |
| 2004 | Muthomi Nyurui |
| 2005 | Tito Sammy / Syengo Robert |
| 2006 | Muneeni Ukulo / Kavula Muthungu |
| 2007 | Kyalo Mumbere / Nancy Maluki |
| 2008 | David Mbiti / Savali Muatha |
| 2009 | Fred Maluki |
| 2010 | Kyambi Kimwele |

## Generating the Excel File

```bash
uv run --with openpyxl python create_excel.py
```

## Contributing

To add or update pupil records:
1. Edit `classes.md` following the existing format (`## Class of YYYY`)
2. Run the script above to regenerate `classes.xlsx`
